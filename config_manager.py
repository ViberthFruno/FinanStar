# Archivo: config_manager.py
# Ubicación: raíz del proyecto
# Descripción: Gestiona la configuración y almacenamiento en JSON con soporte para casos dinámicos

import json
import os
from typing import Dict, List


class ConfigManager:
    DEFAULT_POSITIVE_DEBIT_CODES: Dict[str, str] = {
        'DP': 'DEP',
        'TF': 'T/D',
        'WD': 'T/D',
        '3V': 'O/D',
        '3Y': 'TEF',
        'PE': 'T/D',
        'MD': 'T/D',
        'PT': 'T/D',
    }

    DEFAULT_NON_NEGATIVE_CREDIT_CODES: Dict[str, str] = {
        'DP': 'DEP',
        'AR': 'DEP',
        'TF': 'T/C',
        'MC': 'T/C',
        'PP': 'T/C',
        'WC': 'T/C',
    }

    DEFAULT_DESCRIPTION_OVERRIDES: List[Dict[str, str]] = [
        {'search_text': 'PENDIENTE EN CAMARA DCD', 'code': 'O/C'},
    ]

    def __init__(self, config_file="config.json"):
        """Inicializa el gestor de configuración"""
        self.config_file = config_file

        # Nombres de las 4 cuentas compartidas por los casos 3 y 6
        shared_accounts = [
            "VENTAS F.R. UNO S.A.",
            "NARGALLO DEL ESTE S A",
            "SU LAKA CREANDO SOLUCIONES SOC",
            "3-102-726951 SOCIEDAD DE RESPO"
        ]

        self.CASE_ACCOUNTS = {
            'case3': list(shared_accounts),
            'case6': list(shared_accounts),
            'case9': list(shared_accounts),
            'case12': list(shared_accounts),
        }

        # Mantener atributos legacy para compatibilidad externa
        self.CASE3_ACCOUNTS = self.CASE_ACCOUNTS['case3']
        self.CASE6_ACCOUNTS = self.CASE_ACCOUNTS['case6']
        self.CASE9_ACCOUNTS = self.CASE_ACCOUNTS['case9']
        self.CASE12_ACCOUNTS = self.CASE_ACCOUNTS['case12']

    def load_config(self):
        """Carga la configuración desde el archivo JSON"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as file:
                    return json.load(file)
            else:
                return {}
        except Exception as e:
            print(f"Error al cargar la configuración: {str(e)}")
            return {}

    def save_config(self, config):
        """Guarda la configuración en el archivo JSON"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as file:
                json.dump(config, file, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error al guardar la configuración: {str(e)}")
            return False

    def get_value(self, key, default=None):
        """Obtiene un valor específico de la configuración"""
        config = self.load_config()
        return config.get(key, default)

    def set_value(self, key, value):
        """Establece un valor específico en la configuración"""
        config = self.load_config()
        config[key] = value
        return self.save_config(config)

    def get_email_config(self):
        """Obtiene la configuración de correo electrónico"""
        config = self.load_config()
        return {
            'provider': config.get('provider', ''),
            'email': config.get('email', ''),
            'password': config.get('password', '')
        }

    def set_email_config(self, provider, email, password):
        """Establece la configuración de correo electrónico"""
        config = self.load_config()
        config['provider'] = provider
        config['email'] = email
        config['password'] = password
        return self.save_config(config)

    def get_search_params(self):
        """Obtiene todos los parámetros de búsqueda"""
        config = self.load_config()
        return config.get('search_params', {})

    def set_search_params(self, search_params):
        """Establece todos los parámetros de búsqueda"""
        config = self.load_config()
        config['search_params'] = search_params
        return self.save_config(config)

    def get_case_keyword(self, case_name):
        """Obtiene la palabra clave para un caso específico"""
        search_params = self.get_search_params()
        return search_params.get(case_name, '')

    def set_case_keyword(self, case_name, keyword):
        """Establece la palabra clave para un caso específico"""
        config = self.load_config()
        if 'search_params' not in config:
            config['search_params'] = {}

        if keyword.strip():
            config['search_params'][case_name] = keyword.strip()
        else:
            if case_name in config['search_params']:
                del config['search_params'][case_name]

        return self.save_config(config)

    def get_case1_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 1"""
        config = self.load_config()
        filters = config.get('case1_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case1_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 1"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case1_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case2_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 2"""
        config = self.load_config()
        filters = config.get('case2_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case2_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 2"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case2_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case7_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 7"""
        config = self.load_config()
        filters = config.get('case7_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case7_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 7"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case7_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case8_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 8"""
        config = self.load_config()
        filters = config.get('case8_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case8_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 8"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case8_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case10_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 10"""
        config = self.load_config()
        filters = config.get('case10_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case10_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 10"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case10_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case11_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 11"""
        config = self.load_config()
        filters = config.get('case11_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case11_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 11"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case11_filters'] = cleaned_filters
        return self.save_config(config)

    def _get_case_columns_to_remove(self, key: str):
        """Lee de configuración la lista de columnas a eliminar para un caso."""
        config = self.load_config()
        columns = config.get(key, [])
        if isinstance(columns, list):
            return [str(item) for item in columns if isinstance(item, str)]
        return []

    def _set_case_columns_to_remove(self, key: str, columns):
        """Guarda en configuración la lista de columnas a eliminar para un caso."""
        config = self.load_config()
        if not isinstance(columns, list):
            columns = []
        cleaned_columns = [
            item.strip()
            for item in columns
            if isinstance(item, str) and item.strip()
        ]
        config[key] = cleaned_columns
        return self.save_config(config)

    def _get_code_rules_section(self) -> Dict[str, Dict[str, Dict[str, str]]]:
        config = self.load_config()
        code_rules = config.get('code_rules')
        if isinstance(code_rules, dict):
            return code_rules
        return {}

    def _save_code_rules_section(self, code_rules: Dict[str, Dict[str, Dict[str, str]]]) -> bool:
        config = self.load_config()
        config['code_rules'] = code_rules
        return self.save_config(config)

    def get_positive_debit_code_map(self, case_key: str) -> Dict[str, str]:
        code_rules = self._get_code_rules_section()
        category = code_rules.get('positive_debits')
        defaults = dict(self.DEFAULT_POSITIVE_DEBIT_CODES)
        if isinstance(category, dict) and case_key in category:
            stored = category.get(case_key, {})
            if isinstance(stored, dict):
                return {
                    str(key).strip().upper(): str(value).strip().upper()
                    for key, value in stored.items()
                    if str(key).strip() and str(value).strip()
                }
            return {}
        return defaults

    def set_positive_debit_code_map(self, case_key: str, mapping: Dict[str, str]) -> bool:
        cleaned_map = {
            str(key).strip().upper(): str(value).strip().upper()
            for key, value in mapping.items()
            if str(key).strip() and str(value).strip()
        }
        code_rules = self._get_code_rules_section()
        category = code_rules.setdefault('positive_debits', {})
        category[case_key] = cleaned_map
        return self._save_code_rules_section(code_rules)

    def get_non_negative_credit_code_map(self, case_key: str) -> Dict[str, str]:
        code_rules = self._get_code_rules_section()
        category = code_rules.get('non_negative_credits')
        defaults = dict(self.DEFAULT_NON_NEGATIVE_CREDIT_CODES)
        if isinstance(category, dict) and case_key in category:
            stored = category.get(case_key, {})
            if isinstance(stored, dict):
                return {
                    str(key).strip().upper(): str(value).strip().upper()
                    for key, value in stored.items()
                    if str(key).strip() and str(value).strip()
                }
            return {}
        return defaults

    def set_non_negative_credit_code_map(self, case_key: str, mapping: Dict[str, str]) -> bool:
        cleaned_map = {
            str(key).strip().upper(): str(value).strip().upper()
            for key, value in mapping.items()
            if str(key).strip() and str(value).strip()
        }
        code_rules = self._get_code_rules_section()
        category = code_rules.setdefault('non_negative_credits', {})
        category[case_key] = cleaned_map
        return self._save_code_rules_section(code_rules)

    def get_description_override_rules(self, case_key: str) -> List[Dict[str, str]]:
        code_rules = self._get_code_rules_section()
        category = code_rules.get('description_overrides')
        if isinstance(category, dict) and case_key in category:
            stored_rules = category.get(case_key, [])
            cleaned_rules: List[Dict[str, str]] = []
            if isinstance(stored_rules, list):
                for rule in stored_rules:
                    if not isinstance(rule, dict):
                        continue
                    search_text = str(rule.get('search_text', '')).strip()
                    code = str(rule.get('code', '')).strip().upper()
                    cleaned_rules.append({'search_text': search_text, 'code': code})
            return cleaned_rules
        return [
            {'search_text': item['search_text'], 'code': item['code']}
            for item in self.DEFAULT_DESCRIPTION_OVERRIDES
        ]

    def set_description_override_rules(self, case_key: str, rules: List[Dict[str, str]]) -> bool:
        cleaned_rules: List[Dict[str, str]] = []
        for rule in rules:
            if not isinstance(rule, dict):
                continue
            search_text = str(rule.get('search_text', '')).strip()
            code = str(rule.get('code', '')).strip().upper()
            if not search_text or not code:
                continue
            cleaned_rules.append({'search_text': search_text, 'code': code})

        code_rules = self._get_code_rules_section()
        category = code_rules.setdefault('description_overrides', {})
        category[case_key] = cleaned_rules
        return self._save_code_rules_section(code_rules)

    # ==================== MÉTODOS PARA CASOS CON CUENTAS ====================

    def get_case_account_names(self, case_key):
        """Obtiene la lista de nombres de cuentas para un caso específico"""
        accounts = self.CASE_ACCOUNTS.get(case_key, [])
        return list(accounts)

    def get_case_account_config(self, case_key, account_name):
        """Obtiene la configuración completa de una cuenta para un caso específico"""
        accounts = self.CASE_ACCOUNTS.get(case_key)
        if accounts is None or account_name not in accounts:
            return None

        config = self.load_config()
        case_accounts_key = f"{case_key}_accounts"
        case_accounts = config.get(case_accounts_key, {})

        account_config = case_accounts.get(account_name, {})

        return {
            'codes': account_config.get('codes', []),
            'providers': account_config.get('providers', []),
            'subtypes': account_config.get('subtypes', [])
        }

    def set_case_account_config(self, case_key, account_name, account_config):
        """Establece la configuración completa de una cuenta para un caso específico"""
        accounts = self.CASE_ACCOUNTS.get(case_key)
        if accounts is None or account_name not in accounts:
            return False

        if not isinstance(account_config, dict):
            return False

        config = self.load_config()
        case_accounts_key = f"{case_key}_accounts"

        if case_accounts_key not in config:
            config[case_accounts_key] = {}

        # Limpiar y validar codes
        codes = account_config.get('codes', [])
        if not isinstance(codes, list):
            codes = []
        cleaned_codes = [
            code.strip()
            for code in codes
            if isinstance(code, str) and code.strip()
        ]

        # Limpiar y validar providers
        providers = account_config.get('providers', [])
        if not isinstance(providers, list):
            providers = []
        cleaned_providers = []
        for provider in providers:
            if isinstance(provider, dict):
                search_text = provider.get('search_text', '').strip()
                provider_code = provider.get('provider_code', '').strip()
                if search_text and provider_code:
                    cleaned_providers.append({
                        'search_text': search_text,
                        'provider_code': provider_code
                    })

        # Limpiar y validar subtypes
        subtypes = account_config.get('subtypes', [])
        if not isinstance(subtypes, list):
            subtypes = []
        cleaned_subtypes = []
        for subtype in subtypes:
            if isinstance(subtype, dict):
                document_type = subtype.get('document_type', '').strip()
                search_text = subtype.get('search_text', '').strip()
                subtype_value = subtype.get('subtype_value', '').strip()
                if document_type and search_text and subtype_value:
                    cleaned_subtypes.append({
                        'document_type': document_type,
                        'search_text': search_text,
                        'subtype_value': subtype_value
                    })

        config[case_accounts_key][account_name] = {
            'codes': cleaned_codes,
            'providers': cleaned_providers,
            'subtypes': cleaned_subtypes
        }

        return self.save_config(config)

    def get_case3_account_names(self):
        """Obtiene la lista de nombres de cuentas del Caso 3"""
        return self.get_case_account_names('case3')

    def get_case3_account_config(self, account_name):
        """Obtiene la configuración completa de una cuenta específica del Caso 3"""
        return self.get_case_account_config('case3', account_name)

    def set_case3_account_config(self, account_name, account_config):
        """Establece la configuración completa de una cuenta específica del Caso 3"""
        return self.set_case_account_config('case3', account_name, account_config)

    def get_case6_account_names(self):
        """Obtiene la lista de nombres de cuentas del Caso 6"""
        return self.get_case_account_names('case6')

    def get_case6_account_config(self, account_name):
        """Obtiene la configuración completa de una cuenta específica del Caso 6"""
        return self.get_case_account_config('case6', account_name)

    def set_case6_account_config(self, account_name, account_config):
        """Establece la configuración completa de una cuenta específica del Caso 6"""
        return self.set_case_account_config('case6', account_name, account_config)

    def get_case9_account_names(self):
        """Obtiene la lista de nombres de cuentas del Caso 9"""
        return self.get_case_account_names('case9')

    def get_case9_account_config(self, account_name):
        """Obtiene la configuración completa de una cuenta específica del Caso 9"""
        return self.get_case_account_config('case9', account_name)

    def set_case9_account_config(self, account_name, account_config):
        """Establece la configuración completa de una cuenta específica del Caso 9"""
        return self.set_case_account_config('case9', account_name, account_config)

    def get_case12_account_names(self):
        """Obtiene la lista de nombres de cuentas del Caso 12"""
        return self.get_case_account_names('case12')

    def get_case12_account_config(self, account_name):
        """Obtiene la configuración completa de una cuenta específica del Caso 12"""
        return self.get_case_account_config('case12', account_name)

    def set_case12_account_config(self, account_name, account_config):
        """Establece la configuración completa de una cuenta específica del Caso 12"""
        return self.set_case_account_config('case12', account_name, account_config)

    def find_account_by_code(self, code, case_key='case3'):
        """
        Busca y retorna el nombre de la cuenta que contiene el código especificado
        Retorna: nombre de la cuenta o None si no se encuentra
        """
        if not code or not isinstance(code, str):
            return None

        code_clean = code.strip()
        if not code_clean:
            return None

        config = self.load_config()
        case_accounts_key = f"{case_key}_accounts"
        case_accounts = config.get(case_accounts_key, {})

        for account_name in self.get_case_account_names(case_key):
            account_config = case_accounts.get(account_name, {})
            codes = account_config.get('codes', [])

            if code_clean in codes:
                return account_name

        return None

    # ==================== MÉTODOS LEGACY CASO 3 (MANTENER POR COMPATIBILIDAD) ====================

    def get_case3_providers(self):
        """
        LEGACY: Obtiene la lista global de proveedores (mantener por compatibilidad)
        NOTA: Este método quedará obsoleto con la nueva lógica
        """
        config = self.load_config()
        providers = config.get('case3_providers', [])
        if isinstance(providers, list):
            valid_providers = []
            for item in providers:
                if isinstance(item, dict):
                    search_text = item.get('search_text', '')
                    provider_code = item.get('provider_code', '')
                    if isinstance(search_text, str) and isinstance(provider_code, str):
                        if search_text.strip() and provider_code.strip():
                            valid_providers.append({
                                'search_text': search_text.strip(),
                                'provider_code': provider_code.strip()
                            })
            return valid_providers
        return []

    def set_case3_providers(self, providers):
        """
        LEGACY: Almacena la lista global de proveedores (mantener por compatibilidad)
        NOTA: Este método quedará obsoleto con la nueva lógica
        """
        config = self.load_config()
        if not isinstance(providers, list):
            providers = []

        cleaned_providers = []
        for item in providers:
            if isinstance(item, dict):
                search_text = item.get('search_text', '').strip()
                provider_code = item.get('provider_code', '').strip()
                if search_text and provider_code:
                    cleaned_providers.append({
                        'search_text': search_text,
                        'provider_code': provider_code
                    })

        config['case3_providers'] = cleaned_providers
        return self.save_config(config)

    def get_case3_subtypes(self):
        """
        LEGACY: Obtiene la lista global de subtipos (mantener por compatibilidad)
        NOTA: Este método quedará obsoleto con la nueva lógica
        """
        config = self.load_config()
        subtypes = config.get('case3_subtypes', [])
        if isinstance(subtypes, list):
            valid_subtypes = []
            for item in subtypes:
                if isinstance(item, dict):
                    document_type = item.get('document_type', '')
                    search_text = item.get('search_text', '')
                    subtype_value = item.get('subtype_value', '')
                    if isinstance(document_type, str) and isinstance(search_text, str) and isinstance(subtype_value,
                                                                                                      str):
                        if document_type.strip() and search_text.strip() and subtype_value.strip():
                            valid_subtypes.append({
                                'document_type': document_type.strip(),
                                'search_text': search_text.strip(),
                                'subtype_value': subtype_value.strip()
                            })
            return valid_subtypes
        return []

    def set_case3_subtypes(self, subtypes):
        """
        LEGACY: Almacena la lista global de subtipos (mantener por compatibilidad)
        NOTA: Este método quedará obsoleto con la nueva lógica
        """
        config = self.load_config()
        if not isinstance(subtypes, list):
            subtypes = []

        cleaned_subtypes = []
        for item in subtypes:
            if isinstance(item, dict):
                document_type = item.get('document_type', '').strip()
                search_text = item.get('search_text', '').strip()
                subtype_value = item.get('subtype_value', '').strip()
                if document_type and search_text and subtype_value:
                    cleaned_subtypes.append({
                        'document_type': document_type,
                        'search_text': search_text,
                        'subtype_value': subtype_value
                    })

        config['case3_subtypes'] = cleaned_subtypes
        return self.save_config(config)

    # ==================== FIN MÉTODOS LEGACY ====================

    def _get_case_specific_codification_rules(self, storage_key: str):
        config = self.load_config()
        rules = config.get(storage_key, {})

        def _clean_entries(entries):
            cleaned = []
            if not isinstance(entries, list):
                return cleaned
            for item in entries:
                if not isinstance(item, dict):
                    continue
                search_text = item.get('search_text', '')
                code = item.get('code', '')
                if isinstance(search_text, str) and isinstance(code, str):
                    search_text = search_text.strip()
                    code = code.strip()
                    if search_text and code:
                        cleaned.append({'search_text': search_text, 'code': code})
            return cleaned

        return {
            'debit': _clean_entries(rules.get('debit')),
            'credit': _clean_entries(rules.get('credit')),
        }

    def _set_case_specific_codification_rules(self, storage_key: str, rules):
        if not isinstance(rules, dict):
            rules = {}

        def _clean_entries(entries):
            cleaned = []
            if not isinstance(entries, list):
                return cleaned
            for item in entries:
                if not isinstance(item, dict):
                    continue
                search_text = item.get('search_text', '')
                code = item.get('code', '')
                if isinstance(search_text, str) and isinstance(code, str):
                    search_text = search_text.strip()
                    code = code.strip()
                    if search_text and code:
                        cleaned.append({'search_text': search_text, 'code': code})
            return cleaned

        cleaned_rules = {
            'debit': _clean_entries(rules.get('debit')),
            'credit': _clean_entries(rules.get('credit')),
        }

        config = self.load_config()
        config[storage_key] = cleaned_rules
        return self.save_config(config)

    def get_case4_codification_rules(self):
        """Obtiene las reglas de codificación configuradas para el Caso 4."""
        return self._get_case_specific_codification_rules('case4_codification')

    def set_case4_codification_rules(self, rules):
        """Almacena las reglas de codificación configuradas para el Caso 4."""
        return self._set_case_specific_codification_rules('case4_codification', rules)

    def get_case4_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 4."""
        config = self.load_config()
        filters = config.get('case4_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case4_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 4."""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case4_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case5_filters(self):
        """Obtiene la lista de filtros configurados para el Caso 5"""
        config = self.load_config()
        filters = config.get('case5_filters', [])
        if isinstance(filters, list):
            return [str(item) for item in filters if isinstance(item, str)]
        return []

    def set_case5_filters(self, filters):
        """Almacena la lista de filtros configurados para el Caso 5"""
        config = self.load_config()
        if not isinstance(filters, list):
            filters = []
        cleaned_filters = [
            item.strip()
            for item in filters
            if isinstance(item, str) and item.strip()
        ]
        config['case5_filters'] = cleaned_filters
        return self.save_config(config)

    def get_case5_codification_rules(self):
        """Obtiene las reglas de codificación configuradas para el Caso 5."""
        return self._get_case_specific_codification_rules('case5_codification')

    def set_case5_codification_rules(self, rules):
        """Almacena las reglas de codificación configuradas para el Caso 5."""
        return self._set_case_specific_codification_rules('case5_codification', rules)

    def get_case5_columns_to_remove(self):
        """Obtiene la lista de columnas configuradas para eliminar en el Caso 5."""
        return self._get_case_columns_to_remove('case5_columns_to_remove')

    def set_case5_columns_to_remove(self, columns):
        """Almacena la lista de columnas a eliminar configuradas para el Caso 5."""
        return self._set_case_columns_to_remove('case5_columns_to_remove', columns)

    def get_case7_codification_rules(self):
        """Obtiene las reglas de codificación configuradas para el Caso 7."""
        return self._get_case_specific_codification_rules('case7_codification')

    def set_case7_codification_rules(self, rules):
        """Almacena las reglas de codificación configuradas para el Caso 7."""
        return self._set_case_specific_codification_rules('case7_codification', rules)

    def get_case8_codification_rules(self):
        """Obtiene las reglas de codificación configuradas para el Caso 8."""
        return self._get_case_specific_codification_rules('case8_codification')

    def set_case8_codification_rules(self, rules):
        """Almacena las reglas de codificación configuradas para el Caso 8."""
        return self._set_case_specific_codification_rules('case8_codification', rules)

    def get_case8_columns_to_remove(self):
        """Obtiene la lista de columnas configuradas para eliminar en el Caso 8."""
        return self._get_case_columns_to_remove('case8_columns_to_remove')

    def set_case8_columns_to_remove(self, columns):
        """Almacena la lista de columnas a eliminar configuradas para el Caso 8."""
        return self._set_case_columns_to_remove('case8_columns_to_remove', columns)

    def get_case10_codification_rules(self):
        """Obtiene las reglas de codificación configuradas para el Caso 10."""
        return self._get_case_specific_codification_rules('case10_codification')

    def set_case10_codification_rules(self, rules):
        """Almacena las reglas de codificación configuradas para el Caso 10."""
        return self._set_case_specific_codification_rules('case10_codification', rules)

    def get_case11_codification_rules(self):
        """Obtiene las reglas de codificación configuradas para el Caso 11."""
        return self._get_case_specific_codification_rules('case11_codification')

    def set_case11_codification_rules(self, rules):
        """Almacena las reglas de codificación configuradas para el Caso 11."""
        return self._set_case_specific_codification_rules('case11_codification', rules)

    def get_case11_columns_to_remove(self):
        """Obtiene la lista de columnas configuradas para eliminar en el Caso 11."""
        return self._get_case_columns_to_remove('case11_columns_to_remove')

    def set_case11_columns_to_remove(self, columns):
        """Almacena la lista de columnas a eliminar configuradas para el Caso 11."""
        return self._set_case_columns_to_remove('case11_columns_to_remove', columns)

    def remove_case_keyword(self, case_name):
        """Elimina la palabra clave de un caso específico"""
        config = self.load_config()
        if 'search_params' in config and case_name in config['search_params']:
            del config['search_params'][case_name]
            return self.save_config(config)
        return True

    def get_all_case_keywords(self):
        """Obtiene todas las palabras clave configuradas con sus casos"""
        search_params = self.get_search_params()
        return [(case_name, keyword) for case_name, keyword in search_params.items() if keyword.strip()]

    def has_email_config(self):
        """Verifica si existe configuración completa de correo"""
        email_config = self.get_email_config()
        return all([email_config['provider'], email_config['email'], email_config['password']])

    def has_search_params(self):
        """Verifica si existen parámetros de búsqueda configurados"""
        search_params = self.get_search_params()
        return bool(search_params)

    def validate_config(self):
        """Valida la configuración completa"""
        validation_result = {
            'valid': True,
            'errors': [],
            'warnings': []
        }

        if not self.has_email_config():
            validation_result['valid'] = False
            validation_result['errors'].append("Configuración de correo incompleta")

        if not self.has_search_params():
            validation_result['warnings'].append("No hay parámetros de búsqueda configurados")

        try:
            config = self.load_config()
            if not isinstance(config, dict):
                validation_result['valid'] = False
                validation_result['errors'].append("Archivo de configuración corrupto")
        except Exception as e:
            validation_result['valid'] = False
            validation_result['errors'].append(f"Error al validar configuración: {str(e)}")

        return validation_result

    def reset_config(self):
        """Resetea la configuración a valores por defecto"""
        default_config = {
            'provider': '',
            'email': '',
            'password': '',
            'search_params': {}
        }
        return self.save_config(default_config)

    def backup_config(self, backup_file=None):
        """Crea una copia de seguridad de la configuración"""
        if backup_file is None:
            backup_file = f"{self.config_file}.backup"

        try:
            config = self.load_config()
            with open(backup_file, 'w', encoding='utf-8') as file:
                json.dump(config, file, indent=4, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error al crear copia de seguridad: {str(e)}")
            return False

    def restore_config(self, backup_file=None):
        """Restaura la configuración desde una copia de seguridad"""
        if backup_file is None:
            backup_file = f"{self.config_file}.backup"

        try:
            if os.path.exists(backup_file):
                with open(backup_file, 'r', encoding='utf-8') as file:
                    config = json.load(file)
                return self.save_config(config)
            else:
                print(f"Archivo de copia de seguridad no encontrado: {backup_file}")
                return False
        except Exception as e:
            print(f"Error al restaurar configuración: {str(e)}")
            return False