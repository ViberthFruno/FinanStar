# Archivo: ui_manager.py
# Ubicación: raíz del proyecto
# Descripción: Gestiona la estructura y componentes de la interfaz de usuario

import tkinter as tk
from tkinter import messagebox, ttk
import tkinter.font as tkfont
import threading
import time
from typing import Callable, Dict, List, Optional

from email_manager import EmailManager
from config_manager import ConfigManager
from logger import Logger


class UIManager:
    def __init__(self, root):
        """Inicializa la interfaz de usuario del bot"""
        self.root = root
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(family="Arial", size=10)
        self.root.option_add("*Font", default_font)

        self.email_manager = EmailManager()
        self.config_manager = ConfigManager()
        self.logger = Logger()

        self.monitoring = False
        self.monitor_thread = None

        self.setup_main_frame()
        self.setup_top_panel()
        self.setup_bottom_left_panel()
        self.setup_bottom_right_panel()
        self.initialize_components()

    def setup_main_frame(self):
        """Configura el marco principal de la aplicación"""
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(0, weight=2)
        self.main_frame.rowconfigure(1, weight=1)

    def setup_top_panel(self):
        """Configura el panel superior con sistema de pestañas por banco"""
        self.top_panel = ttk.LabelFrame(self.main_frame, text="Panel Principal")
        self.top_panel.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=5, pady=5)

        # Crear el Notebook (sistema de pestañas)
        self.notebook = ttk.Notebook(self.top_panel)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Crear pestañas para cada banco
        self.setup_bac_tab()
        self.setup_davivienda_tab()
        self.setup_promerica_tab()
        self.setup_bcr_tab()

    def setup_bac_tab(self):
        """Configura la pestaña del banco BAC (Casos 1-3)"""
        bac_frame = ttk.Frame(self.notebook)
        self.notebook.add(bac_frame, text="BAC")

        bac_frame.columnconfigure(0, weight=1)
        bac_frame.columnconfigure(1, weight=1)

        # Caso 1
        case1_frame = ttk.LabelFrame(bac_frame, text="📄 Caso 1 - Formato Mejorado", padding="10")
        case1_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case1_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case1_filters_modal
        ).pack(fill=tk.X, pady=5)
        self._add_code_rule_buttons(case1_frame, 'case1', 'Caso 1')

        # Caso 2
        case2_frame = ttk.LabelFrame(bac_frame, text="📄 Caso 2 - Filtro por Fechas", padding="10")
        case2_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case2_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case2_filters_modal
        ).pack(fill=tk.X, pady=5)
        self._add_code_rule_buttons(case2_frame, 'case2', 'Caso 2')

        # Caso 3
        case3_frame = ttk.LabelFrame(bac_frame, text="📄 Caso 3 - Plantillas Estándar por Cuenta", padding="10")
        case3_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        account_names = self.config_manager.get_case3_account_names()

        for idx, account_name in enumerate(account_names):
            row = idx // 2
            col = idx % 2

            btn = ttk.Button(
                case3_frame,
                text=f"⚙ {account_name}",
                command=lambda name=account_name: self.open_case3_account_modal(name),
                width=35
            )
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        case3_frame.columnconfigure(0, weight=1)
        case3_frame.columnconfigure(1, weight=1)

    def setup_davivienda_tab(self):
        """Configura la pestaña del banco Davivienda (Casos 4-6)"""
        davivienda_frame = ttk.Frame(self.notebook)
        self.notebook.add(davivienda_frame, text="Davivienda")

        davivienda_frame.columnconfigure(0, weight=1)
        davivienda_frame.columnconfigure(1, weight=1)

        # Caso 4
        case4_frame = ttk.LabelFrame(davivienda_frame, text="📄 Caso 4 - Rediseño con Codificación", padding="10")
        case4_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case4_frame,
            text="⚙ Configurar Codificación",
            command=self.open_case4_codification_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case4_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case4_filters_modal
        ).pack(fill=tk.X, pady=5)

        # Caso 5
        case5_frame = ttk.LabelFrame(davivienda_frame, text="📄 Caso 5 - Formato con Filtro", padding="10")
        case5_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case5_frame,
            text="⚙ Configurar Codificación",
            command=self.open_case5_codification_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case5_frame,
            text="⚙ Configurar Columnas a Eliminar",
            command=self.open_case5_column_removal_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case5_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case5_filters_modal
        ).pack(fill=tk.X, pady=5)

        # Caso 6
        case6_frame = ttk.LabelFrame(davivienda_frame, text="📄 Caso 6 - Plantillas Estándar por Cuenta", padding="10")
        case6_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        case6_account_names = self.config_manager.get_case6_account_names()

        for idx, account_name in enumerate(case6_account_names):
            row = idx // 2
            col = idx % 2

            btn = ttk.Button(
                case6_frame,
                text=f"⚙ {account_name}",
                command=lambda name=account_name: self.open_case6_account_modal(name),
                width=35
            )
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        case6_frame.columnconfigure(0, weight=1)
        case6_frame.columnconfigure(1, weight=1)

    def setup_promerica_tab(self):
        """Configura la pestaña del banco Promerica (Casos 7-9)"""
        promerica_frame = ttk.Frame(self.notebook)
        self.notebook.add(promerica_frame, text="Promerica")

        promerica_frame.columnconfigure(0, weight=1)
        promerica_frame.columnconfigure(1, weight=1)

        # Caso 7
        case7_frame = ttk.LabelFrame(promerica_frame, text="📄 Caso 7 - Formato Verde", padding="10")
        case7_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case7_frame,
            text="⚙ Configurar Codificación",
            command=self.open_case7_codification_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case7_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case7_filters_modal
        ).pack(fill=tk.X, pady=5)

        # Caso 8
        case8_frame = ttk.LabelFrame(promerica_frame, text="📄 Caso 8 - Formato Verde sin Débitos", padding="10")
        case8_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case8_frame,
            text="⚙ Configurar Codificación",
            command=self.open_case8_codification_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case8_frame,
            text="⚙ Configurar Columnas a Eliminar",
            command=self.open_case8_column_removal_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case8_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case8_filters_modal
        ).pack(fill=tk.X, pady=5)

        # Caso 9
        case9_frame = ttk.LabelFrame(promerica_frame, text="📄 Caso 9 - Plantillas desde Caso 7", padding="10")
        case9_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        case9_account_names = self.config_manager.get_case9_account_names()

        for idx, account_name in enumerate(case9_account_names):
            row = idx // 2
            col = idx % 2

            btn = ttk.Button(
                case9_frame,
                text=f"⚙ {account_name}",
                command=lambda name=account_name: self.open_case9_account_modal(name),
                width=35
            )
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        case9_frame.columnconfigure(0, weight=1)
        case9_frame.columnconfigure(1, weight=1)

    def setup_bcr_tab(self):
        """Configura la pestaña del banco BCR (Casos 10-12)"""
        bcr_frame = ttk.Frame(self.notebook)
        self.notebook.add(bcr_frame, text="BCR")

        bcr_frame.columnconfigure(0, weight=1)
        bcr_frame.columnconfigure(1, weight=1)

        # Caso 10
        case10_frame = ttk.LabelFrame(bcr_frame, text="📄 Caso 10 - Formato Celeste", padding="10")
        case10_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case10_frame,
            text="⚙ Configurar Codificación",
            command=self.open_case10_codification_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case10_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case10_filters_modal
        ).pack(fill=tk.X, pady=5)

        # Caso 11
        case11_frame = ttk.LabelFrame(bcr_frame, text="📄 Caso 11 - Formato Celeste sin Débitos", padding="10")
        case11_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=5)
        ttk.Button(
            case11_frame,
            text="⚙ Configurar Codificación",
            command=self.open_case11_codification_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case11_frame,
            text="⚙ Configurar Columnas a Eliminar",
            command=self.open_case11_column_removal_modal
        ).pack(fill=tk.X, pady=5)
        ttk.Button(
            case11_frame,
            text="⚙ Configurar Filtrados",
            command=self.open_case11_filters_modal
        ).pack(fill=tk.X, pady=5)

        # Caso 12
        case12_frame = ttk.LabelFrame(bcr_frame, text="📄 Caso 12 - Plantillas desde Caso 10", padding="10")
        case12_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=5)

        case12_account_names = self.config_manager.get_case12_account_names()

        for idx, account_name in enumerate(case12_account_names):
            row = idx // 2
            col = idx % 2

            btn = ttk.Button(
                case12_frame,
                text=f"⚙ {account_name}",
                command=lambda name=account_name: self.open_case12_account_modal(name),
                width=35
            )
            btn.grid(row=row, column=col, padx=5, pady=5, sticky="ew")

        case12_frame.columnconfigure(0, weight=1)
        case12_frame.columnconfigure(1, weight=1)

    def _add_code_rule_buttons(self, parent_frame, case_key: str, case_display_name: str) -> None:
        """Agrega botones comunes para configurar reglas de codificación."""
        button_specs = [
            (
                "🔁 Códigos de Débito",
                lambda: self.open_code_mapping_modal(case_key, case_display_name, 'positive_debits'),
            ),
            (
                "💳 Códigos de Crédito",
                lambda: self.open_code_mapping_modal(case_key, case_display_name, 'non_negative_credits'),
            ),
            (
                "📝 Reglas por Descripción",
                lambda: self.open_description_override_modal(case_key, case_display_name),
            ),
        ]

        for text, command in button_specs:
            ttk.Button(parent_frame, text=text, command=command).pack(fill=tk.X, pady=2)

    def _open_codification_rules_modal(
            self,
            case_display_name: str,
            window_title: str,
            description: str,
            debit_section_title: str,
            credit_section_title: str,
            get_rules: Callable[[], Dict[str, List[Dict[str, str]]]],
            set_rules: Callable[[Dict[str, List[Dict[str, str]]]], bool],
            success_message: str,
            error_message: str,
    ) -> None:
        """Muestra un modal reutilizable para configurar reglas de codificación por descripción."""
        rules = get_rules() or {}
        debit_rules = [dict(item) for item in rules.get('debit', [])]
        credit_rules = [dict(item) for item in rules.get('credit', [])]

        modal = tk.Toplevel(self.root)
        modal.title(window_title)
        modal.geometry("700x520")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        main_frame = ttk.Frame(modal, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            main_frame,
            text=description,
            wraplength=640,
            justify=tk.LEFT,
        ).pack(anchor="w", padx=5, pady=(0, 10))

        def build_rules_section(section_title: str, rules_list: List[Dict[str, str]]):
            section_frame = ttk.LabelFrame(main_frame, text=section_title, padding="10")
            section_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

            list_frame = ttk.Frame(section_frame)
            list_frame.pack(fill=tk.BOTH, expand=True)

            scrollbar = ttk.Scrollbar(list_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
            listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=listbox.yview)

            def refresh_listbox() -> None:
                listbox.delete(0, tk.END)
                for rule in rules_list:
                    search_text = rule.get('search_text', '')
                    code = rule.get('code', '')
                    listbox.insert(tk.END, f"Buscar: '{search_text}' → Código: '{code}'")

            refresh_listbox()

            form_frame = ttk.Frame(section_frame)
            form_frame.pack(fill=tk.X, pady=10)

            ttk.Label(form_frame, text="Texto a buscar:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
            search_var = tk.StringVar()
            search_entry = ttk.Entry(form_frame, textvariable=search_var)
            search_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

            ttk.Label(form_frame, text="Código a asignar:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
            code_var = tk.StringVar()
            code_entry = ttk.Entry(form_frame, textvariable=code_var)
            code_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

            def add_rule() -> None:
                search_text = search_var.get().strip()
                code_text = code_var.get().strip()

                if not search_text or not code_text:
                    self.logger.log(
                        "Debe completar tanto el texto a buscar como el código para agregar una regla.",
                        level="WARNING",
                    )
                    return

                rules_list.append({'search_text': search_text, 'code': code_text})
                refresh_listbox()
                search_var.set('')
                code_var.set('')
                self.logger.log(
                    f"[{case_display_name}] Regla agregada: '{search_text}' → '{code_text}'",
                    level="INFO",
                )

            def remove_rule() -> None:
                selection = listbox.curselection()
                if not selection:
                    self.logger.log("Seleccione una regla para eliminar.", level="WARNING")
                    return

                index = selection[0]
                removed = rules_list.pop(index)
                refresh_listbox()
                self.logger.log(
                    f"[{case_display_name}] Regla eliminada: '{removed.get('search_text', '')}' → '{removed.get('code', '')}'",
                    level="INFO",
                )

            def on_select(_event) -> None:
                selection = listbox.curselection()
                if not selection:
                    return
                rule = rules_list[selection[0]]
                search_var.set(rule.get('search_text', ''))
                code_var.set(rule.get('code', ''))

            listbox.bind('<<ListboxSelect>>', on_select)

            button_frame = ttk.Frame(form_frame)
            button_frame.grid(row=2, column=0, columnspan=2, pady=10)

            ttk.Button(button_frame, text="Agregar", command=add_rule).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="Eliminar Seleccionada", command=remove_rule).pack(side=tk.LEFT, padx=5)

            form_frame.columnconfigure(1, weight=1)

        build_rules_section(debit_section_title, debit_rules)
        build_rules_section(credit_section_title, credit_rules)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_rules() -> None:
            payload = {
                'debit': debit_rules,
                'credit': credit_rules,
            }

            if set_rules(payload):
                self.logger.log(success_message, level="INFO")
                modal.destroy()
            else:
                self.logger.log(error_message, level="ERROR")
                messagebox.showerror("Error", error_message)

        ttk.Button(button_frame, text="Guardar", command=save_rules).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=5
        )
        ttk.Button(button_frame, text="Cancelar", command=modal.destroy).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=5
        )

    def open_case_account_modal(self, case_key, case_display_name, account_name):
        """Abre un modal de configuración completa para una cuenta de un caso específico"""
        account_config = self.config_manager.get_case_account_config(case_key, account_name)

        if account_config is None:
            self.logger.log(
                f"Error: cuenta '{account_name}' no válida en {case_display_name}",
                level="ERROR"
            )
            return

        modal = tk.Toplevel(self.root)
        modal.title(f"{case_display_name} - Configuración de {account_name}")
        modal.geometry("900x700")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        main_frame = ttk.Frame(modal, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        codes_section = ttk.LabelFrame(main_frame, text="Códigos de Cuenta", padding="10")
        codes_section.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Label(
            codes_section,
            text="Ingrese los códigos que identifican archivos de esta cuenta (uno por línea):",
            wraplength=850
        ).pack(anchor="w", padx=5, pady=(0, 5))

        codes_text = tk.Text(codes_section, wrap=tk.WORD, height=4)
        codes_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        codes_text.insert(tk.END, "\n".join(account_config['codes']))

        providers_section = ttk.LabelFrame(main_frame, text="Proveedores", padding="10")
        providers_section.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Label(
            providers_section,
            text="Configure los proveedores que se asignarán automáticamente para esta cuenta:",
            wraplength=850
        ).pack(anchor="w", padx=5, pady=(0, 5))

        providers_list_frame = ttk.Frame(providers_section)
        providers_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        providers_scrollbar = ttk.Scrollbar(providers_list_frame)
        providers_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        providers_listbox = tk.Listbox(providers_list_frame, yscrollcommand=providers_scrollbar.set, height=4)
        providers_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        providers_scrollbar.config(command=providers_listbox.yview)

        providers_data = list(account_config['providers'])

        def refresh_providers_listbox():
            providers_listbox.delete(0, tk.END)
            for provider in providers_data:
                search_text = provider.get('search_text', '')
                provider_code = provider.get('provider_code', '')
                providers_listbox.insert(tk.END, f"Buscar: '{search_text}' → Proveedor: '{provider_code}'")

        refresh_providers_listbox()

        providers_form = ttk.Frame(providers_section)
        providers_form.pack(fill=tk.X, pady=5)

        ttk.Label(providers_form, text="Texto a buscar:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        provider_search_var = tk.StringVar()
        ttk.Entry(providers_form, textvariable=provider_search_var, width=30).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=5,
            pady=2
        )

        ttk.Label(providers_form, text="Código de proveedor:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        provider_code_var = tk.StringVar()
        ttk.Entry(providers_form, textvariable=provider_code_var, width=30).grid(
            row=0,
            column=3,
            sticky="ew",
            padx=5,
            pady=2
        )

        def add_provider():
            search_text = provider_search_var.get().strip()
            provider_code = provider_code_var.get().strip()
            if not search_text or not provider_code:
                self.logger.log("Debe completar ambos campos del proveedor", level="WARNING")
                return
            providers_data.append({'search_text': search_text, 'provider_code': provider_code})
            refresh_providers_listbox()
            provider_search_var.set("")
            provider_code_var.set("")
            self.logger.log(
                f"[{case_display_name}] Proveedor agregado: '{search_text}' → '{provider_code}'",
                level="INFO"
            )

        def remove_provider():
            selection = providers_listbox.curselection()
            if not selection:
                self.logger.log("Seleccione un proveedor para eliminar", level="WARNING")
                return
            index = selection[0]
            removed = providers_data.pop(index)
            refresh_providers_listbox()
            self.logger.log(
                f"[{case_display_name}] Proveedor eliminado: '{removed['search_text']}'",
                level="INFO"
            )

        providers_buttons = ttk.Frame(providers_form)
        providers_buttons.grid(row=1, column=0, columnspan=4, pady=5)
        ttk.Button(providers_buttons, text="Agregar", command=add_provider).pack(side=tk.LEFT, padx=5)
        ttk.Button(providers_buttons, text="Eliminar Seleccionado", command=remove_provider).pack(side=tk.LEFT, padx=5)

        providers_form.columnconfigure(1, weight=1)
        providers_form.columnconfigure(3, weight=1)

        subtypes_section = ttk.LabelFrame(main_frame, text="Sub Tipos de Documento", padding="10")
        subtypes_section.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        ttk.Label(
            subtypes_section,
            text="Configure los subtipos de documento para esta cuenta:",
            wraplength=850
        ).pack(anchor="w", padx=5, pady=(0, 5))

        subtypes_list_frame = ttk.Frame(subtypes_section)
        subtypes_list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        subtypes_scrollbar = ttk.Scrollbar(subtypes_list_frame)
        subtypes_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        subtypes_listbox = tk.Listbox(subtypes_list_frame, yscrollcommand=subtypes_scrollbar.set, height=4)
        subtypes_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        subtypes_scrollbar.config(command=subtypes_listbox.yview)

        subtypes_data = list(account_config['subtypes'])

        def refresh_subtypes_listbox():
            subtypes_listbox.delete(0, tk.END)
            for subtype in subtypes_data:
                document_type = subtype.get('document_type', '')
                search_text = subtype.get('search_text', '')
                subtype_value = subtype.get('subtype_value', '')
                subtypes_listbox.insert(
                    tk.END,
                    f"Tipo: '{document_type}' + Texto: '{search_text}' → Subtipo: '{subtype_value}'"
                )

        refresh_subtypes_listbox()

        subtypes_form = ttk.Frame(subtypes_section)
        subtypes_form.pack(fill=tk.X, pady=5)

        ttk.Label(subtypes_form, text="Tipo documento:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        subtype_doctype_var = tk.StringVar()
        ttk.Entry(subtypes_form, textvariable=subtype_doctype_var, width=15).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=5,
            pady=2
        )

        ttk.Label(subtypes_form, text="Texto en concepto:").grid(row=0, column=2, sticky="w", padx=5, pady=2)
        subtype_search_var = tk.StringVar()
        ttk.Entry(subtypes_form, textvariable=subtype_search_var, width=20).grid(
            row=0,
            column=3,
            sticky="ew",
            padx=5,
            pady=2
        )

        ttk.Label(subtypes_form, text="Valor subtipo:").grid(row=0, column=4, sticky="w", padx=5, pady=2)
        subtype_value_var = tk.StringVar()
        ttk.Entry(subtypes_form, textvariable=subtype_value_var, width=15).grid(
            row=0,
            column=5,
            sticky="ew",
            padx=5,
            pady=2
        )

        def add_subtype():
            document_type = subtype_doctype_var.get().strip()
            search_text = subtype_search_var.get().strip()
            subtype_value = subtype_value_var.get().strip()
            if not document_type or not search_text or not subtype_value:
                self.logger.log("Debe completar los tres campos del subtipo", level="WARNING")
                return
            subtypes_data.append({
                'document_type': document_type,
                'search_text': search_text,
                'subtype_value': subtype_value
            })
            refresh_subtypes_listbox()
            subtype_doctype_var.set("")
            subtype_search_var.set("")
            subtype_value_var.set("")
            self.logger.log(
                f"[{case_display_name}] Subtipo agregado: Tipo '{document_type}' + Texto '{search_text}'",
                level="INFO"
            )

        def remove_subtype():
            selection = subtypes_listbox.curselection()
            if not selection:
                self.logger.log("Seleccione un subtipo para eliminar", level="WARNING")
                return
            index = selection[0]
            removed = subtypes_data.pop(index)
            refresh_subtypes_listbox()
            self.logger.log(
                f"[{case_display_name}] Subtipo eliminado: Tipo '{removed['document_type']}'",
                level="INFO"
            )

        subtypes_buttons = ttk.Frame(subtypes_form)
        subtypes_buttons.grid(row=1, column=0, columnspan=6, pady=5)
        ttk.Button(subtypes_buttons, text="Agregar", command=add_subtype).pack(side=tk.LEFT, padx=5)
        ttk.Button(subtypes_buttons, text="Eliminar Seleccionado", command=remove_subtype).pack(side=tk.LEFT, padx=5)

        subtypes_form.columnconfigure(1, weight=1)
        subtypes_form.columnconfigure(3, weight=1)
        subtypes_form.columnconfigure(5, weight=1)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_all():
            codes_text_value = codes_text.get("1.0", tk.END)
            codes = [
                code.strip()
                for code in codes_text_value.splitlines()
                if code.strip()
            ]

            full_config = {
                'codes': codes,
                'providers': providers_data,
                'subtypes': subtypes_data
            }

            if self.config_manager.set_case_account_config(case_key, account_name, full_config):
                self.logger.log(
                    f"Configuración de '{account_name}' guardada correctamente en {case_display_name}",
                    level="INFO"
                )
                modal.destroy()
            else:
                self.logger.log(
                    f"Error al guardar configuración de '{account_name}' en {case_display_name}",
                    level="ERROR"
                )

        save_button = ttk.Button(button_frame, text="Guardar Configuración", command=save_all)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case3_account_modal(self, account_name):
        """Abre el modal de configuración para una cuenta del Caso 3"""
        self.open_case_account_modal('case3', 'Caso 3', account_name)

    def open_case6_account_modal(self, account_name):
        """Abre el modal de configuración para una cuenta del Caso 6"""
        self.open_case_account_modal('case6', 'Caso 6', account_name)

    def open_case9_account_modal(self, account_name):
        """Abre el modal de configuración para una cuenta del Caso 9"""
        self.open_case_account_modal('case9', 'Caso 9', account_name)

    def open_case12_account_modal(self, account_name):
        """Abre el modal de configuración para una cuenta del Caso 12"""
        self.open_case_account_modal('case12', 'Caso 12', account_name)

    def setup_bottom_left_panel(self):
        """Configura el panel inferior izquierdo para configuración de correo"""
        self.bottom_left_panel = ttk.LabelFrame(self.main_frame, text="Configuración de Correo")
        self.bottom_left_panel.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        self.config_button = ttk.Button(
            self.bottom_left_panel,
            text="Configurar Correo",
            command=self.open_email_config_modal
        )
        self.config_button.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        self.search_params_button = ttk.Button(
            self.bottom_left_panel,
            text="Parametros de Busqueda",
            command=self.open_search_params_modal
        )
        self.search_params_button.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        self.cc_users_button = ttk.Button(
            self.bottom_left_panel,
            text="Usuarios Adjuntos (CC)",
            command=self.open_cc_users_modal
        )
        self.cc_users_button.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        monitor_frame = ttk.LabelFrame(self.bottom_left_panel, text="⚙ Control de Monitoreo", padding="15")
        monitor_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        button_container = ttk.Frame(monitor_frame)
        button_container.pack()

        self.monitor_button = ttk.Button(
            button_container,
            text="▶ Iniciar Monitoreo",
            command=self.toggle_monitoring,
            width=30
        )
        self.monitor_button.pack(side=tk.LEFT, padx=5)

        self.status_label = ttk.Label(
            button_container,
            text="● Detenido",
            foreground="red",
            font=("Arial", 10, "bold")
        )
        self.status_label.pack(side=tk.LEFT, padx=10)

        self.bottom_left_panel.columnconfigure(0, weight=1)

    def open_cc_users_modal(self):
        """Abre una ventana modal para configurar correos en CC"""
        config = self.config_manager.load_config()
        cc_users_list = config.get('cc_users', [])

        modal = tk.Toplevel(self.root)
        modal.title("Configurar Usuarios Adjuntos (CC)")
        modal.geometry("400x300")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        cc_frame = ttk.Frame(modal, padding="10")
        cc_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(cc_frame, text="Ingrese los correos a copiar (uno por línea):").pack(anchor="w", padx=5, pady=(0, 5))

        cc_text = tk.Text(cc_frame, wrap=tk.WORD, height=10)
        cc_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        cc_text.insert(tk.END, "\n".join(cc_users_list))

        button_frame = ttk.Frame(cc_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_cc_users():
            emails_text = cc_text.get("1.0", tk.END).strip()
            emails_list = [email.strip() for email in emails_text.split("\n") if email.strip()]

            current_config = self.config_manager.load_config()
            current_config['cc_users'] = emails_list

            if self.config_manager.save_config(current_config):
                self.logger.log("Lista de usuarios CC guardada correctamente.", level="INFO")
                modal.destroy()
            else:
                self.logger.log("Error al guardar la lista de usuarios CC.", level="ERROR")

        save_button = ttk.Button(button_frame, text="Guardar", command=save_cc_users)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case1_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 1."""
        filters_list = self.config_manager.get_case1_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 1")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case1_filters(filters):
                self.logger.log(
                    "Filtros del Caso 1 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 1.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case2_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 2."""
        filters_list = self.config_manager.get_case2_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 2")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case2_filters(filters):
                self.logger.log(
                    "Filtros del Caso 2 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 2.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case7_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 7."""
        filters_list = self.config_manager.get_case7_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 7")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case7_filters(filters):
                self.logger.log(
                    "Filtros del Caso 7 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 7.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case8_column_removal_modal(self):
        """Abre el modal para configurar columnas a eliminar del Caso 8."""
        self._open_column_removal_modal(
            title="Columnas a Eliminar - Caso 8",
            instructions=(
                "Indique el encabezado exacto de las columnas que no deben aparecer en el reporte. "
                "Escriba una columna por línea."
            ),
            get_columns=self.config_manager.get_case8_columns_to_remove,
            set_columns=self.config_manager.set_case8_columns_to_remove,
            success_message="Columnas a eliminar del Caso 8 guardadas correctamente.",
            error_message="Error al guardar las columnas a eliminar del Caso 8.",
        )

    def open_case8_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 8."""
        filters_list = self.config_manager.get_case8_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 8")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case8_filters(filters):
                self.logger.log(
                    "Filtros del Caso 8 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 8.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_code_mapping_modal(self, case_key: str, case_display_name: str, mapping_type: str) -> None:
        """Abre un modal genérico para configurar mapeos de códigos por monto."""
        if mapping_type == 'positive_debits':
            title = f"{case_display_name} - Códigos de Débito"
            description = (
                "Defina los códigos que se sustituirán cuando el monto del débito sea "
                "mayor a cero y no exista crédito en la misma fila."
            )
            mapping_data = dict(self.config_manager.get_positive_debit_code_map(case_key))
            save_action = self.config_manager.set_positive_debit_code_map
            success_message = f"Reglas de códigos de débito para {case_display_name} guardadas correctamente."
            error_message = f"Error al guardar las reglas de códigos de débito para {case_display_name}."
        else:
            title = f"{case_display_name} - Códigos de Crédito"
            description = (
                "Defina los códigos que se sustituirán cuando el monto del crédito sea "
                "mayor a cero y no exista débito en la misma fila."
            )
            mapping_data = dict(self.config_manager.get_non_negative_credit_code_map(case_key))
            save_action = self.config_manager.set_non_negative_credit_code_map
            success_message = f"Reglas de códigos de crédito para {case_display_name} guardadas correctamente."
            error_message = f"Error al guardar las reglas de códigos de crédito para {case_display_name}."

        modal = tk.Toplevel(self.root)
        modal.title(title)
        modal.geometry("500x440")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        container = ttk.Frame(modal, padding="12")
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            container,
            text=description,
            wraplength=440,
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 8))

        list_frame = ttk.Frame(container)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(list_frame, height=8, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        displayed_codes: List[str] = []

        def refresh_listbox() -> None:
            nonlocal displayed_codes
            displayed_codes = sorted(mapping_data.keys())
            listbox.delete(0, tk.END)
            for code in displayed_codes:
                listbox.insert(tk.END, f"{code} → {mapping_data[code]}")

        form_frame = ttk.Frame(container)
        form_frame.pack(fill=tk.X, pady=10)
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Código original:").grid(row=0, column=0, sticky="w")
        original_entry = ttk.Entry(form_frame)
        original_entry.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        ttk.Label(form_frame, text="Código nuevo:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        new_entry = ttk.Entry(form_frame)
        new_entry.grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))

        def reset_entries() -> None:
            original_entry.delete(0, tk.END)
            new_entry.delete(0, tk.END)

        def add_or_update() -> None:
            original_code = original_entry.get().strip().upper()
            new_code = new_entry.get().strip().upper()

            if not original_code or not new_code:
                messagebox.showerror("Entrada inválida", "Debe indicar el código original y el código nuevo.")
                return

            mapping_data[original_code] = new_code
            refresh_listbox()
            listbox.selection_clear(0, tk.END)
            selection_index = displayed_codes.index(original_code) if original_code in displayed_codes else None
            if selection_index is not None:
                listbox.selection_set(selection_index)
                listbox.see(selection_index)
            reset_entries()

        def delete_selected() -> None:
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Sin selección", "Seleccione un código para eliminar.")
                return
            code = displayed_codes[selection[0]]
            mapping_data.pop(code, None)
            refresh_listbox()
            reset_entries()

        def on_select(_event) -> None:
            selection = listbox.curselection()
            if not selection:
                return
            code = displayed_codes[selection[0]]
            original_entry.delete(0, tk.END)
            original_entry.insert(0, code)
            new_entry.delete(0, tk.END)
            new_entry.insert(0, mapping_data.get(code, ''))

        listbox.bind('<<ListboxSelect>>', on_select)

        action_frame = ttk.Frame(container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(action_frame, text="Agregar / Actualizar", command=add_or_update).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5)
        )
        ttk.Button(action_frame, text="Eliminar Seleccionado", command=delete_selected).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0)
        )

        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X)

        def save_changes() -> None:
            if save_action(case_key, mapping_data):
                self.logger.log(success_message, level="INFO")
                modal.destroy()
            else:
                messagebox.showerror("Error", error_message)

        ttk.Button(button_frame, text="Guardar", command=save_changes).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5)
        )
        ttk.Button(button_frame, text="Cancelar", command=modal.destroy).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0)
        )

        refresh_listbox()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

    def open_description_override_modal(self, case_key: str, case_display_name: str) -> None:
        """Abre un modal para configurar reglas por coincidencia en la descripción."""
        rules = [
            {
                'search_text': str(rule.get('search_text', '')).strip(),
                'code': str(rule.get('code', '')).strip().upper(),
            }
            for rule in self.config_manager.get_description_override_rules(case_key)
        ]

        modal = tk.Toplevel(self.root)
        modal.title(f"{case_display_name} - Reglas por Descripción")
        modal.geometry("520x480")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        container = ttk.Frame(modal, padding="12")
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            container,
            text=(
                "Especifique los textos a buscar en la columna Descripción y el código "
                "que se asignará cuando exista coincidencia. La búsqueda no distingue "
                "mayúsculas, minúsculas ni acentos."
            ),
            wraplength=460,
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 8))

        list_frame = ttk.Frame(container)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        listbox = tk.Listbox(list_frame, height=8, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)

        form_frame = ttk.Frame(container)
        form_frame.pack(fill=tk.X, pady=10)
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="Texto a buscar:").grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(form_frame)
        search_entry.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        ttk.Label(form_frame, text="Código a asignar:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        code_entry = ttk.Entry(form_frame)
        code_entry.grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=(6, 0))

        selected_index: Optional[int] = None

        def refresh_rules() -> None:
            listbox.delete(0, tk.END)
            for rule in rules:
                listbox.insert(
                    tk.END,
                    f"Buscar: \"{rule['search_text']}\" → Código: \"{rule['code']}\"",
                )

        def reset_entries() -> None:
            search_entry.delete(0, tk.END)
            code_entry.delete(0, tk.END)

        def add_or_update_rule() -> None:
            nonlocal selected_index
            search_text = search_entry.get().strip()
            new_code = code_entry.get().strip().upper()

            if not search_text or not new_code:
                messagebox.showerror(
                    "Entrada inválida",
                    "Debe indicar el texto a buscar y el código a asignar.",
                )
                return

            if selected_index is not None and 0 <= selected_index < len(rules):
                rules[selected_index] = {'search_text': search_text, 'code': new_code}
            else:
                normalized_search = search_text.lower()
                for idx, rule in enumerate(rules):
                    if rule['search_text'].lower() == normalized_search:
                        rules[idx] = {'search_text': search_text, 'code': new_code}
                        selected_index = idx
                        break
                else:
                    rules.append({'search_text': search_text, 'code': new_code})
                    selected_index = len(rules) - 1

            refresh_rules()
            if selected_index is not None:
                listbox.selection_clear(0, tk.END)
                listbox.selection_set(selected_index)
                listbox.see(selected_index)
            reset_entries()

        def delete_rule() -> None:
            nonlocal selected_index
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Sin selección", "Seleccione una regla para eliminar.")
                return
            index = selection[0]
            if 0 <= index < len(rules):
                rules.pop(index)
                selected_index = None
                refresh_rules()
                reset_entries()

        def on_rule_select(_event) -> None:
            nonlocal selected_index
            selection = listbox.curselection()
            if not selection:
                selected_index = None
                return
            selected_index = selection[0]
            if 0 <= selected_index < len(rules):
                rule = rules[selected_index]
                search_entry.delete(0, tk.END)
                search_entry.insert(0, rule['search_text'])
                code_entry.delete(0, tk.END)
                code_entry.insert(0, rule['code'])

        listbox.bind('<<ListboxSelect>>', on_rule_select)

        action_frame = ttk.Frame(container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(action_frame, text="Agregar / Actualizar", command=add_or_update_rule).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5)
        )
        ttk.Button(action_frame, text="Eliminar Seleccionado", command=delete_rule).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0)
        )

        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X)

        def save_rules() -> None:
            if self.config_manager.set_description_override_rules(case_key, rules):
                self.logger.log(
                    f"Reglas por descripción para {case_display_name} guardadas correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                messagebox.showerror(
                    "Error",
                    f"No fue posible guardar las reglas por descripción para {case_display_name}.",
                )

        ttk.Button(button_frame, text="Guardar", command=save_rules).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5)
        )
        ttk.Button(button_frame, text="Cancelar", command=modal.destroy).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0)
        )

        refresh_rules()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

    def open_case10_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 10."""
        filters_list = self.config_manager.get_case10_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 10")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case10_filters(filters):
                self.logger.log(
                    "Filtros del Caso 10 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 10.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case10_codification_modal(self):
        """Abre un modal para configurar las reglas de codificación del Caso 10."""
        self._open_codification_rules_modal(
            case_display_name="Caso 10",
            window_title="Reglas de Codificación del Caso 10",
            description=(
                "Defina las palabras clave que se buscarán en la columna Descripción para asignar "
                "automáticamente un código en la columna Código."
            ),
            debit_section_title="Reglas para Débitos",
            credit_section_title="Reglas para Créditos",
            get_rules=self.config_manager.get_case10_codification_rules,
            set_rules=self.config_manager.set_case10_codification_rules,
            success_message="Reglas de codificación del Caso 10 guardadas correctamente.",
            error_message="Error al guardar las reglas de codificación del Caso 10.",
        )

    def open_case7_codification_modal(self):
        """Abre un modal para configurar las reglas de codificación del Caso 7."""
        self._open_codification_rules_modal(
            case_display_name="Caso 7",
            window_title="Reglas de Codificación del Caso 7",
            description=(
                "Configure las palabras clave que completarán la columna Código del formato verde "
                "del Caso 7 según la descripción de cada movimiento."
            ),
            debit_section_title="Reglas para Débitos",
            credit_section_title="Reglas para Créditos",
            get_rules=self.config_manager.get_case7_codification_rules,
            set_rules=self.config_manager.set_case7_codification_rules,
            success_message="Reglas de codificación del Caso 7 guardadas correctamente.",
            error_message="Error al guardar las reglas de codificación del Caso 7.",
        )

    def open_case8_codification_modal(self):
        """Abre un modal para configurar las reglas de codificación del Caso 8."""
        self._open_codification_rules_modal(
            case_display_name="Caso 8",
            window_title="Reglas de Codificación del Caso 8",
            description=(
                "Defina las palabras clave que completarán la columna Código del formato sin "
                "débitos del Caso 8."
            ),
            debit_section_title="Reglas para Débitos",
            credit_section_title="Reglas para Créditos",
            get_rules=self.config_manager.get_case8_codification_rules,
            set_rules=self.config_manager.set_case8_codification_rules,
            success_message="Reglas de codificación del Caso 8 guardadas correctamente.",
            error_message="Error al guardar las reglas de codificación del Caso 8.",
        )

    def open_case11_codification_modal(self):
        """Abre un modal para configurar las reglas de codificación del Caso 11."""
        self._open_codification_rules_modal(
            case_display_name="Caso 11",
            window_title="Reglas de Codificación del Caso 11",
            description=(
                "Especifique las palabras clave que llenarán la columna Código en el formato "
                "celeste del Caso 11, sin columna de débitos."
            ),
            debit_section_title="Reglas para Débitos",
            credit_section_title="Reglas para Créditos",
            get_rules=self.config_manager.get_case11_codification_rules,
            set_rules=self.config_manager.set_case11_codification_rules,
            success_message="Reglas de codificación del Caso 11 guardadas correctamente.",
            error_message="Error al guardar las reglas de codificación del Caso 11.",
        )

    def open_case11_column_removal_modal(self):
        """Abre el modal para configurar columnas a eliminar del Caso 11."""
        self._open_column_removal_modal(
            title="Columnas a Eliminar - Caso 11",
            instructions=(
                "Liste los encabezados de columna que deben ocultarse en el archivo final. "
                "Escriba una columna por línea."
            ),
            get_columns=self.config_manager.get_case11_columns_to_remove,
            set_columns=self.config_manager.set_case11_columns_to_remove,
            success_message="Columnas a eliminar del Caso 11 guardadas correctamente.",
            error_message="Error al guardar las columnas a eliminar del Caso 11.",
        )

    def open_case11_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 11."""
        filters_list = self.config_manager.get_case11_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 11")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case11_filters(filters):
                self.logger.log(
                    "Filtros del Caso 11 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 11.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case4_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 4."""
        filters_list = self.config_manager.get_case4_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 4")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada palabra u oración a buscar en la columna Descripción (una por línea).\n"
                "Las filas coincidentes se resaltarán y marcarán en la columna Revisar."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case4_filters(filters):
                self.logger.log(
                    "Filtros del Caso 4 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 4.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def open_case4_codification_modal(self):
        """Abre un modal para configurar las reglas de codificación del Caso 4."""
        self._open_codification_rules_modal(
            case_display_name="Caso 4",
            window_title="Reglas de Codificación del Caso 4",
            description=(
                "Defina las palabras clave que se buscarán en la columna Descripción para asignar "
                "automáticamente un código en la columna Código."
            ),
            debit_section_title="Reglas para Débitos (DR)",
            credit_section_title="Reglas para Créditos (CR)",
            get_rules=self.config_manager.get_case4_codification_rules,
            set_rules=self.config_manager.set_case4_codification_rules,
            success_message="Reglas de codificación del Caso 4 guardadas correctamente.",
            error_message="Error al guardar las reglas de codificación del Caso 4.",
        )

    def open_case5_codification_modal(self):
        """Abre un modal para configurar las reglas de codificación del Caso 5."""
        self._open_codification_rules_modal(
            case_display_name="Caso 5",
            window_title="Reglas de Codificación del Caso 5",
            description=(
                "Indique las palabras clave que se buscarán en la descripción para completar la "
                "columna Código del formato filtrado del Caso 5."
            ),
            debit_section_title="Reglas para Débitos (DR)",
            credit_section_title="Reglas para Créditos (CR)",
            get_rules=self.config_manager.get_case5_codification_rules,
            set_rules=self.config_manager.set_case5_codification_rules,
            success_message="Reglas de codificación del Caso 5 guardadas correctamente.",
            error_message="Error al guardar las reglas de codificación del Caso 5.",
        )

    def open_case5_column_removal_modal(self):
        """Abre el modal para configurar columnas a eliminar del Caso 5."""
        self._open_column_removal_modal(
            title="Columnas a Eliminar - Caso 5",
            instructions=(
                "Ingrese el nombre del encabezado de cada columna exactamente como aparece en el "
                "archivo final. Escriba una columna por línea."
            ),
            get_columns=self.config_manager.get_case5_columns_to_remove,
            set_columns=self.config_manager.set_case5_columns_to_remove,
            success_message="Columnas a eliminar del Caso 5 guardadas correctamente.",
            error_message="Error al guardar las columnas a eliminar del Caso 5.",
        )

    def open_case5_filters_modal(self):
        """Abre un modal para configurar las oraciones de filtrado del Caso 5."""
        filters_list = self.config_manager.get_case5_filters()

        modal = tk.Toplevel(self.root)
        modal.title("Filtrados del Caso 5")
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        filters_frame = ttk.Frame(modal, padding="10")
        filters_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            filters_frame,
            text=(
                "Ingrese cada oración a buscar en la columna Descripción (una por línea).\n"
                "Las coincidencias se resaltarán automáticamente en amarillo y marcarán la columna Revisar."
            ),
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        filters_text = tk.Text(filters_frame, wrap=tk.WORD, height=12)
        filters_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        filters_text.insert(tk.END, "\n".join(filters_list))

        button_frame = ttk.Frame(filters_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_filters():
            filters_text_value = filters_text.get("1.0", tk.END)
            filters = [
                sentence.strip()
                for sentence in filters_text_value.splitlines()
                if sentence.strip()
            ]

            if self.config_manager.set_case5_filters(filters):
                self.logger.log(
                    "Filtros del Caso 5 guardados correctamente.",
                    level="INFO",
                )
                modal.destroy()
            else:
                self.logger.log(
                    "Error al guardar los filtros del Caso 5.",
                    level="ERROR",
                )

        save_button = ttk.Button(button_frame, text="Guardar", command=save_filters)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

    def _open_column_removal_modal(
            self,
            title: str,
            instructions: str,
            get_columns: Callable[[], List[str]],
            set_columns: Callable[[List[str]], bool],
            success_message: str,
            error_message: str,
    ) -> None:
        columns_list = get_columns()

        modal = tk.Toplevel(self.root)
        modal.title(title)
        modal.geometry("420x320")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        container = ttk.Frame(modal, padding="10")
        container.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            container,
            text=instructions,
            justify=tk.LEFT,
            wraplength=360,
        ).pack(anchor="w", padx=5, pady=(0, 5))

        columns_text = tk.Text(container, wrap=tk.WORD, height=12)
        columns_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        columns_text.insert(tk.END, "\n".join(columns_list))

        button_frame = ttk.Frame(container)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_columns():
            text_value = columns_text.get("1.0", tk.END)
            columns = [
                line.strip()
                for line in text_value.splitlines()
                if line.strip()
            ]

            if set_columns(columns):
                self.logger.log(success_message, level="INFO")
                modal.destroy()
            else:
                self.logger.log(error_message, level="ERROR")

        ttk.Button(button_frame, text="Guardar", command=save_columns).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=5
        )
        ttk.Button(button_frame, text="Cancelar", command=modal.destroy).pack(
            side=tk.LEFT, expand=True, fill=tk.X, padx=5
        )

    def open_email_config_modal(self):
        """Abre una ventana modal para la configuración de correo"""
        config = self.config_manager.load_config()

        modal = tk.Toplevel(self.root)
        modal.title("Configuración de Correo")
        modal.geometry("400x250")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        config_frame = ttk.Frame(modal, padding="10")
        config_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(config_frame, text="Proveedor de Correo:").grid(row=0, column=0, sticky="w", padx=5, pady=5)

        provider_var = tk.StringVar(value=config.get('provider', 'Gmail'))
        provider_combo = ttk.Combobox(config_frame, textvariable=provider_var)
        provider_combo['values'] = ('Gmail', 'Outlook', 'Yahoo', 'Otro')
        provider_combo.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(config_frame, text="Usuario (email):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        email_var = tk.StringVar(value=config.get('email', ''))
        email_entry = ttk.Entry(config_frame, textvariable=email_var)
        email_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(config_frame, text="Contraseña:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        password_var = tk.StringVar(value=config.get('password', ''))
        password_entry = ttk.Entry(config_frame, textvariable=password_var, show="*")
        password_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        button_frame = ttk.Frame(config_frame)
        button_frame.grid(row=4, column=0, columnspan=2, sticky="ew", pady=20)

        def test_connection_modal():
            provider = provider_var.get()
            email = email_var.get()
            password = password_var.get()

            if not all([provider, email, password]):
                self.logger.log("Error: Todos los campos son obligatorios", level="ERROR")
                return

            smtp_result = self.email_manager.test_smtp_connection(provider, email, password)
            imap_result = self.email_manager.test_imap_connection(provider, email, password)

            if smtp_result and imap_result:
                self.logger.log(f"Conexión exitosa a {provider} (SMTP e IMAP)", level="INFO")
            else:
                if not smtp_result:
                    self.logger.log(f"Error en la conexión SMTP a {provider}", level="ERROR")
                if not imap_result:
                    self.logger.log(f"Error en la conexión IMAP a {provider}", level="ERROR")

        def save_config_modal():
            current_config = self.config_manager.load_config()
            current_config.update({
                'provider': provider_var.get(),
                'email': email_var.get(),
                'password': password_var.get()
            })

            if not all([current_config['provider'], current_config['email'], current_config['password']]):
                self.logger.log("Error: Todos los campos son obligatorios para guardar", level="ERROR")
                return

            if self.config_manager.save_config(current_config):
                self.logger.log("Configuración guardada correctamente", level="INFO")
                modal.destroy()
            else:
                self.logger.log("Error al guardar la configuración", level="ERROR")

        test_button = ttk.Button(button_frame, text="Probar Conexión", command=test_connection_modal)
        test_button.grid(row=0, column=0, sticky="ew", padx=5)

        save_button = ttk.Button(button_frame, text="Guardar Datos", command=save_config_modal)
        save_button.grid(row=0, column=1, sticky="ew", padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.grid(row=0, column=2, sticky="ew", padx=5)

        config_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)
        button_frame.columnconfigure(2, weight=1)

    def open_search_params_modal(self):
        """Abre una ventana modal para configurar parámetros de búsqueda"""
        config = self.config_manager.load_config()
        search_params = config.get('search_params', {
            'caso1': '',
            'caso2': '',
            'caso3': '',
            'caso4': '',
            'caso5': '',
            'caso6': '',
            'caso7': '',
            'caso8': '',
            'caso9': '',
            'caso10': '',
            'caso11': '',
            'caso12': '',
        })

        modal = tk.Toplevel(self.root)
        modal.title("Parametros de Busqueda")
        modal.geometry("400x420")
        modal.transient(self.root)
        modal.grab_set()
        modal.focus_set()

        modal.update_idletasks()
        width = modal.winfo_width()
        height = modal.winfo_height()
        x = (modal.winfo_screenwidth() // 2) - (width // 2)
        y = (modal.winfo_screenheight() // 2) - (height // 2)
        modal.geometry(f"{width}x{height}+{x}+{y}")

        params_frame = ttk.Frame(modal, padding="10")
        params_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(params_frame, text="Caso 1:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        caso1_var = tk.StringVar(value=search_params.get('caso1', ''))
        caso1_entry = ttk.Entry(params_frame, textvariable=caso1_var)
        caso1_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 2:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        caso2_var = tk.StringVar(value=search_params.get('caso2', ''))
        caso2_entry = ttk.Entry(params_frame, textvariable=caso2_var)
        caso2_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 3:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        caso3_var = tk.StringVar(value=search_params.get('caso3', ''))
        caso3_entry = ttk.Entry(params_frame, textvariable=caso3_var)
        caso3_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 4:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        caso4_var = tk.StringVar(value=search_params.get('caso4', ''))
        caso4_entry = ttk.Entry(params_frame, textvariable=caso4_var)
        caso4_entry.grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 5:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        caso5_var = tk.StringVar(value=search_params.get('caso5', ''))
        caso5_entry = ttk.Entry(params_frame, textvariable=caso5_var)
        caso5_entry.grid(row=4, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 6:").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        caso6_var = tk.StringVar(value=search_params.get('caso6', ''))
        caso6_entry = ttk.Entry(params_frame, textvariable=caso6_var)
        caso6_entry.grid(row=5, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 7:").grid(row=6, column=0, sticky="w", padx=5, pady=5)
        caso7_var = tk.StringVar(value=search_params.get('caso7', ''))
        caso7_entry = ttk.Entry(params_frame, textvariable=caso7_var)
        caso7_entry.grid(row=6, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 8:").grid(row=7, column=0, sticky="w", padx=5, pady=5)
        caso8_var = tk.StringVar(value=search_params.get('caso8', ''))
        caso8_entry = ttk.Entry(params_frame, textvariable=caso8_var)
        caso8_entry.grid(row=7, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 9:").grid(row=8, column=0, sticky="w", padx=5, pady=5)
        caso9_var = tk.StringVar(value=search_params.get('caso9', ''))
        caso9_entry = ttk.Entry(params_frame, textvariable=caso9_var)
        caso9_entry.grid(row=8, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 10:").grid(row=9, column=0, sticky="w", padx=5, pady=5)
        caso10_var = tk.StringVar(value=search_params.get('caso10', ''))
        caso10_entry = ttk.Entry(params_frame, textvariable=caso10_var)
        caso10_entry.grid(row=9, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 11:").grid(row=10, column=0, sticky="w", padx=5, pady=5)
        caso11_var = tk.StringVar(value=search_params.get('caso11', ''))
        caso11_entry = ttk.Entry(params_frame, textvariable=caso11_var)
        caso11_entry.grid(row=10, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(params_frame, text="Caso 12:").grid(row=11, column=0, sticky="w", padx=5, pady=5)
        caso12_var = tk.StringVar(value=search_params.get('caso12', ''))
        caso12_entry = ttk.Entry(params_frame, textvariable=caso12_var)
        caso12_entry.grid(row=11, column=1, sticky="ew", padx=5, pady=5)

        button_frame = ttk.Frame(params_frame)
        button_frame.grid(row=12, column=0, columnspan=2, sticky="ew", pady=20)

        def save_search_params():
            current_config = self.config_manager.load_config()

            current_config['search_params'] = {
                'caso1': caso1_var.get().strip(),
                'caso2': caso2_var.get().strip(),
                'caso3': caso3_var.get().strip(),
                'caso4': caso4_var.get().strip(),
                'caso5': caso5_var.get().strip(),
                'caso6': caso6_var.get().strip(),
                'caso7': caso7_var.get().strip(),
                'caso8': caso8_var.get().strip(),
                'caso9': caso9_var.get().strip(),
                'caso10': caso10_var.get().strip(),
                'caso11': caso11_var.get().strip(),
                'caso12': caso12_var.get().strip(),
            }

            if self.config_manager.save_config(current_config):
                self.logger.log("Parámetros de búsqueda guardados correctamente", level="INFO")
                modal.destroy()
            else:
                self.logger.log("Error al guardar parámetros de búsqueda", level="ERROR")

        save_button = ttk.Button(button_frame, text="Guardar", command=save_search_params)
        save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        cancel_button = ttk.Button(button_frame, text="Cancelar", command=modal.destroy)
        cancel_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        params_frame.columnconfigure(1, weight=1)

    def toggle_monitoring(self):
        """Inicia o detiene el monitoreo de emails"""
        if not self.monitoring:
            config = self.config_manager.load_config()
            if not all([config.get('provider'), config.get('email'), config.get('password')]):
                self.logger.log("Error: Configure primero los datos de correo", level="ERROR")
                return

            search_params = config.get('search_params', {})
            if not any([
                search_params.get('caso1', '').strip(),
                search_params.get('caso2', '').strip(),
                search_params.get('caso3', '').strip(),
                search_params.get('caso4', '').strip(),
                search_params.get('caso5', '').strip(),
                search_params.get('caso6', '').strip(),
                search_params.get('caso7', '').strip(),
                search_params.get('caso8', '').strip(),
                search_params.get('caso9', '').strip(),
                search_params.get('caso10', '').strip(),
                search_params.get('caso11', '').strip(),
            ]):
                self.logger.log("Error: Configure primero los parámetros de búsqueda", level="ERROR")
                return

            self.monitoring = True
            self.monitor_button.config(text="⏸ Detener Monitoreo")
            self.status_label.config(text="● Monitoreando", foreground="green")

            self.monitor_thread = threading.Thread(target=self.monitor_emails, daemon=True)
            self.monitor_thread.start()

            self.logger.log("Monitoreo de emails iniciado", level="INFO")
        else:
            self.monitoring = False
            self.monitor_button.config(text="▶ Iniciar Monitoreo")
            self.status_label.config(text="● Detenido", foreground="red")
            self.logger.log("Monitoreo de emails detenido", level="INFO")

    def monitor_emails(self):
        """Función que se ejecuta en un hilo separado para monitorear emails"""
        while self.monitoring:
            try:
                config = self.config_manager.load_config()
                search_params = config.get('search_params', {})
                cc_list = config.get('cc_users', [])

                search_titles = []
                if search_params.get('caso1', '').strip():
                    search_titles.append(search_params['caso1'].strip())
                if search_params.get('caso2', '').strip():
                    search_titles.append(search_params['caso2'].strip())
                if search_params.get('caso3', '').strip():
                    search_titles.append(search_params['caso3'].strip())
                if search_params.get('caso4', '').strip():
                    search_titles.append(search_params['caso4'].strip())
                if search_params.get('caso5', '').strip():
                    search_titles.append(search_params['caso5'].strip())
                if search_params.get('caso6', '').strip():
                    search_titles.append(search_params['caso6'].strip())
                if search_params.get('caso7', '').strip():
                    search_titles.append(search_params['caso7'].strip())
                if search_params.get('caso8', '').strip():
                    search_titles.append(search_params['caso8'].strip())
                if search_params.get('caso9', '').strip():
                    search_titles.append(search_params['caso9'].strip())
                if search_params.get('caso10', '').strip():
                    search_titles.append(search_params['caso10'].strip())
                if search_params.get('caso11', '').strip():
                    search_titles.append(search_params['caso11'].strip())
                if search_params.get('caso12', '').strip():
                    search_titles.append(search_params['caso12'].strip())

                if search_titles:
                    self.email_manager.check_and_process_emails(
                        config['provider'],
                        config['email'],
                        config['password'],
                        search_titles,
                        self.logger,
                        cc_list
                    )

                time.sleep(30)

            except Exception as e:
                self.logger.log(f"Error en el monitoreo: {str(e)}", level="ERROR")
                time.sleep(60)

    def setup_bottom_right_panel(self):
        """Configura el panel inferior derecho para logs"""
        self.bottom_right_panel = ttk.LabelFrame(self.main_frame, text="Log del Sistema")
        self.bottom_right_panel.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)

        self.log_text = tk.Text(self.bottom_right_panel, wrap=tk.WORD, height=10, width=40)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        scrollbar = ttk.Scrollbar(self.bottom_right_panel, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)

        self.logger.set_text_widget(self.log_text)

    def initialize_components(self):
        """Inicializa componentes adicionales y carga la configuración"""
        config = self.config_manager.load_config()
        if config:
            self.logger.log("Configuración cargada correctamente.")