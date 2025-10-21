"""
LogLife Operations - Refactored Modern UI
A modern, clean interface for LogLife automation operations
"""

import customtkinter as ctk
import datetime as dt
import os
import tempfile
from pathlib import Path
from tkinter import StringVar, IntVar
from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
import tkinter.ttk as ttk

# Import operations and utilities
from operations import *
from utils.functions import Start, Browse

# Modern theme configuration
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Constants
WINDOW_WIDTH = 900
WINDOW_HEIGHT = 650
SIDEBAR_WIDTH = 200
CONTENT_PADDING = 20
COMPONENT_SPACING = 15

# Color scheme
COLORS = {
    'primary': '#1f538d',
    'secondary': '#14375e',
    'accent': '#00a5e7',
    'success': '#28a745',
    'warning': '#ffc107',
    'danger': '#dc3545',
    'light_gray': '#f8f9fa',
    'medium_gray': '#6c757d',
    'dark_gray': '#343a40'
}

# Get system paths
download_path = str(Path.home() / "Downloads")
temp_folder = f'{tempfile.gettempdir()}/Operação LogLife'

# Calendar configuration
today = dt.datetime.today()
cal_config = {
    'selectmode': 'day',
    'day': today.day,
    'month': today.month,
    'year': today.year,
    'locale': 'pt_BR',
    'firstweekday': 'sunday',
    'showweeknumbers': False,
    'bordercolor': "white",
    'background': "white",
    'disabledbackground': "white",
    'headersbackground': "white",
    'normalbackground': "white",
    'normalforeground': 'black',
    'headersforeground': 'black',
    'selectbackground': COLORS['accent'],
    'selectforeground': 'white',
    'weekendbackground': 'white',
    'weekendforeground': 'black',
    'othermonthforeground': 'black',
    'othermonthbackground': '#E8E8E8',
    'othermonthweforeground': 'black',
    'othermonthwebackground': '#E8E8E8',
    'foreground': "black"
}


class ModernCard(ctk.CTkFrame):
    """A modern card component with shadow effect"""
    
    def __init__(self, parent, title=None, **kwargs):
        super().__init__(parent, corner_radius=12, **kwargs)
        
        if title:
            self.title_label = ctk.CTkLabel(
                self, 
                text=title, 
                font=ctk.CTkFont(size=16, weight="bold"),
                text_color=COLORS['primary']
            )
            self.title_label.pack(pady=(15, 10), padx=20, anchor="w")


class ModernButton(ctk.CTkButton):
    """Enhanced button with consistent styling"""
    
    def __init__(self, parent, style="primary", **kwargs):
        # Set default styling based on style type
        if style == "primary":
            color = COLORS['primary']
            hover_color = COLORS['secondary']
        elif style == "success":
            color = COLORS['success']
            hover_color = "#218838"
        elif style == "warning":
            color = COLORS['warning']
            hover_color = "#e0a800"
        elif style == "danger":
            color = COLORS['danger']
            hover_color = "#c82333"
        else:
            color = COLORS['medium_gray']
            hover_color = COLORS['dark_gray']
        
        super().__init__(
            parent,
            fg_color=color,
            hover_color=hover_color,
            corner_radius=8,
            height=36,
            font=ctk.CTkFont(size=13, weight="bold"),
            **kwargs
        )


class StatusIndicator(ctk.CTkFrame):
    """Status indicator with color-coded states"""
    
    def __init__(self, parent, status="idle"):
        super().__init__(parent, corner_radius=20, width=20, height=20)
        
        self.status_colors = {
            'idle': COLORS['medium_gray'],
            'processing': COLORS['warning'],
            'success': COLORS['success'],
            'error': COLORS['danger']
        }
        
        self.configure(fg_color=self.status_colors.get(status, COLORS['medium_gray']))
    
    def update_status(self, status):
        self.configure(fg_color=self.status_colors.get(status, COLORS['medium_gray']))


class FileSelector(ctk.CTkFrame):
    """Modern file selector component"""
    
    def __init__(self, parent, label_text, file_var, browse_handler, **kwargs):
        super().__init__(parent, corner_radius=8, **kwargs)
        
        # Label
        self.label = ctk.CTkLabel(
            self, 
            text=label_text,
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.label.pack(pady=(10, 5), padx=15, anchor="w")
        
        # File path display
        self.file_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.file_frame.pack(fill="x", padx=15, pady=(0, 5))
        
        self.file_display = ctk.CTkLabel(
            self.file_frame,
            text=file_var.get() if file_var.get() else "Nenhum arquivo selecionado",
            font=ctk.CTkFont(size=11),
            text_color=COLORS['medium_gray'],
            wraplength=300,
            anchor="w"
        )
        self.file_display.pack(side="left", fill="x", expand=True)
        
        # Browse button
        self.browse_btn = ModernButton(
            self.file_frame,
            text="📁 Procurar",
            width=100,
            command=browse_handler
        )
        self.browse_btn.pack(side="right", padx=(10, 0))
    
    def update_display(self, text):
        self.file_display.configure(text=text)


class LogLifeModernApp(ctk.CTk):
    """Modern LogLife Operations Application"""
    
    def __init__(self):
        super().__init__()
        
        # Configure main window
        self.title('LogLife Operations - Sistema Moderno')
        self.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}")
        self.resizable(True, True)
        
        # Set icon
        icon_path = Path(__file__).parent.parent / 'assets' / 'my_icon.ico'
        if icon_path.exists():
            self.iconbitmap(str(icon_path))
        
        # Initialize threads
        self.thread_0 = Start(self)
        self.thread_1 = Start(self)
        self.thread_2 = Start(self)
        
        # Create main layout
        self.setup_layout()
        self.setup_variables()
        self.create_interface()
        
    def setup_layout(self):
        """Setup the main layout structure"""
        # Configure grid weights
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Sidebar
        self.sidebar = ctk.CTkFrame(self, width=SIDEBAR_WIDTH, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_rowconfigure(8, weight=1)
        
        # Main content area
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=CONTENT_PADDING, pady=CONTENT_PADDING)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        
    def setup_variables(self):
        """Initialize all variables"""
        # Create temp directories
        os.makedirs(f'{temp_folder}/filepaths', exist_ok=True)
        
        # File variables
        self.filename = StringVar()
        self.filename2 = StringVar()
        self.fleury_sheet_name = StringVar()
        self.folderpath = StringVar()
        self.downloadpath = StringVar()
        
        # Load saved file paths
        self.load_file_paths()
        
        # Form variables
        self.dispatch_prot = StringVar()
        self.material_type = IntVar(value=0)
        self.flight_service = IntVar(value=0)
        self.vols = IntVar(value=1)
        self.kg_record = IntVar(value=1)
        self.date_filter = IntVar(value=1)
        self.reg_filter = IntVar(value=0)
        self.reg = StringVar(value="1")
        
    def load_file_paths(self):
        """Load saved file paths from temp files"""
        file_mappings = {
            'sheet_name.txt': (self.filename, 'Selecione a Planilha de Serviços'),
            'filename2.txt': (self.filename2, 'Selecione a Relação de Cargas'),
            'fleury_sheet_name.txt': (self.fleury_sheet_name, 'Selecione a Planilha Fleury'),
            'folderpath.txt': (self.folderpath, 'Selecione a pasta das minutas'),
            'downloadpath.txt': (self.downloadpath, download_path)
        }
        
        for filename, (var, default) in file_mappings.items():
            try:
                with open(f'{temp_folder}/filepaths/{filename}', 'r') as f:
                    var.set(f.read().strip())
            except FileNotFoundError:
                var.set(default)
    
    def create_interface(self):
        """Create the main interface"""
        self.create_sidebar()
        self.create_header()
        self.create_main_content()
        
    def create_sidebar(self):
        """Create the sidebar navigation"""
        # Logo/Title
        self.logo_label = ctk.CTkLabel(
            self.sidebar,
            text="LogLife\nOperations",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=COLORS['primary']
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 30))
        
        # Navigation buttons
        nav_buttons = [
            ("📊 Serviços", self.show_services),
            ("📋 Minutas", self.show_minutas),
            ("📦 Cargas", self.show_cargas),
            ("📁 Arquivos", self.show_arquivos)
        ]
        
        self.nav_buttons = []
        for i, (text, command) in enumerate(nav_buttons):
            btn = ModernButton(
                self.sidebar,
                text=text,
                width=160,
                command=command,
                anchor="w"
            )
            btn.grid(row=i+1, column=0, padx=20, pady=(0, 10), sticky="ew")
            self.nav_buttons.append(btn)
        
        # Status indicators
        self.status_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.status_frame.grid(row=7, column=0, padx=20, pady=20, sticky="ew")
        
        ctk.CTkLabel(
            self.status_frame,
            text="Status do Sistema:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", pady=(0, 10))
        
        # Individual status indicators
        status_items = [
            ("Serviços", "services_status"),
            ("Minutas", "minutas_status"),
            ("Cargas", "cargas_status")
        ]
        
        self.status_indicators = {}
        for text, key in status_items:
            frame = ctk.CTkFrame(self.status_frame, fg_color="transparent")
            frame.pack(fill="x", pady=2)
            
            indicator = StatusIndicator(frame)
            indicator.pack(side="left")
            
            label = ctk.CTkLabel(
                frame,
                text=text,
                font=ctk.CTkFont(size=11)
            )
            label.pack(side="left", padx=(8, 0))
            
            self.status_indicators[key] = indicator
        
        # Current active section
        self.current_section = "services"
        self.update_nav_selection()
    
    def create_header(self):
        """Create the header section"""
        self.header = ctk.CTkFrame(self.main_frame, height=60, corner_radius=12)
        self.header.grid(row=0, column=0, sticky="ew", pady=(0, CONTENT_PADDING))
        self.header.grid_columnconfigure(1, weight=1)
        
        # Page title
        self.page_title = ctk.CTkLabel(
            self.header,
            text="Gerenciamento de Serviços",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=COLORS['primary']
        )
        self.page_title.grid(row=0, column=0, padx=20, pady=15, sticky="w")
        
        # Quick actions
        self.quick_actions = ctk.CTkFrame(self.header, fg_color="transparent")
        self.quick_actions.grid(row=0, column=1, padx=20, pady=10, sticky="e")
        
    def create_main_content(self):
        """Create the main content area"""
        self.content = ctk.CTkScrollableFrame(self.main_frame, corner_radius=12)
        self.content.grid(row=1, column=0, sticky="nsew")
        self.content.grid_columnconfigure(0, weight=1)
        
        # Create all sections (initially hidden)
        self.create_services_section()
        self.create_minutas_section()
        self.create_cargas_section()
        self.create_arquivos_section()
        
        # Show default section
        self.show_services()
    
    def create_services_section(self):
        """Create the services management section"""
        self.services_frame = ModernCard(self.content, title="Gerenciamento de Serviços")
        
        # Date selection
        date_card = ModernCard(self.services_frame, title="Seleção de Período")
        date_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        date_grid = ctk.CTkFrame(date_card, fg_color="transparent")
        date_grid.pack(fill="x", padx=20, pady=(0, 15))
        date_grid.grid_columnconfigure((0, 1), weight=1)
        
        # Start date
        start_frame = ctk.CTkFrame(date_grid, fg_color="transparent")
        start_frame.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        
        ctk.CTkLabel(
            start_frame,
            text="Data Inicial:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", pady=(0, 5))
        
        self.cal_start = DateEntry(
            start_frame,
            style='my.DateEntry',
            **cal_config
        )
        self.cal_start.pack(fill="x")
        
        # End date
        end_frame = ctk.CTkFrame(date_grid, fg_color="transparent")
        end_frame.grid(row=0, column=1, padx=(10, 0), sticky="ew")
        
        ctk.CTkLabel(
            end_frame,
            text="Data Final:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", pady=(0, 5))
        
        self.cal_end = DateEntry(
            end_frame,
            style='my.DateEntry',
            **cal_config
        )
        self.cal_end.pack(fill="x")
        
        # Filters
        filters_card = ModernCard(self.services_frame, title="Filtros")
        filters_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        filters_grid = ctk.CTkFrame(filters_card, fg_color="transparent")
        filters_grid.pack(fill="x", padx=20, pady=(0, 15))
        filters_grid.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Date filter checkbox
        self.date_filter_chk = ctk.CTkCheckBox(
            filters_grid,
            text='Filtrar por Data',
            variable=self.date_filter
        )
        self.date_filter_chk.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Regional filter checkbox
        self.reg_filter_chk = ctk.CTkCheckBox(
            filters_grid,
            text='Filtrar por Regional',
            variable=self.reg_filter
        )
        self.reg_filter_chk.grid(row=0, column=1, padx=10, pady=5, sticky="w")
        
        # Regional selection
        reg_frame = ctk.CTkFrame(filters_grid, fg_color="transparent")
        reg_frame.grid(row=0, column=2, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(
            reg_frame,
            text="Regional:",
            font=ctk.CTkFont(size=12)
        ).pack(side="left")
        
        self.reg_option = ctk.CTkOptionMenu(
            reg_frame,
            variable=self.reg,
            values=["1", "2", "3", "4", "5", "6"],
            width=70
        )
        self.reg_option.pack(side="right")
        
        # Actions
        actions_card = ModernCard(self.services_frame, title="Ações")
        actions_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        actions_grid = ctk.CTkFrame(actions_card, fg_color="transparent")
        actions_grid.pack(fill="x", padx=20, pady=(0, 15))
        actions_grid.grid_columnconfigure((0, 1), weight=1)
        
        self.update_data_btn = ModernButton(
            actions_grid,
            text="🔄 Atualizar Dados",
            style="primary",
            command=self.update_services_data
        )
        self.update_data_btn.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        
        self.clear_data_btn = ModernButton(
            actions_grid,
            text="🗑️ Limpar Dados",
            style="warning",
            command=self.clear_services_data
        )
        self.clear_data_btn.grid(row=0, column=1, padx=(10, 0), sticky="ew")
        
        # Progress bar
        self.services_progress = ctk.CTkProgressBar(
            self.services_frame,
            mode='indeterminate'
        )
        self.services_progress.pack(fill="x", padx=20, pady=(10, 20))
        
    def create_minutas_section(self):
        """Create the minutas management section"""
        self.minutas_frame = ModernCard(self.content, title="Emissão de Minutas")
        
        # Single minuta
        single_card = ModernCard(self.minutas_frame, title="Minuta Individual")
        single_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        # Protocol input
        prot_frame = ctk.CTkFrame(single_card, fg_color="transparent")
        prot_frame.pack(fill="x", padx=20, pady=(0, 15))
        
        ctk.CTkLabel(
            prot_frame,
            text="Protocolo:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(side="left")
        
        self.protocol_entry = ctk.CTkEntry(
            prot_frame,
            textvariable=self.dispatch_prot,
            width=200,
            placeholder_text="Digite o protocolo"
        )
        self.protocol_entry.pack(side="left", padx=(10, 0))
        
        self.emit_single_btn = ModernButton(
            prot_frame,
            text="📄 Emitir Minuta",
            style="success",
            command=self.emit_single_minuta
        )
        self.emit_single_btn.pack(side="right")
        
        # Configuration options
        config_card = ModernCard(self.minutas_frame, title="Configurações")
        config_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        config_grid = ctk.CTkFrame(config_card, fg_color="transparent")
        config_grid.pack(fill="x", padx=20, pady=(0, 15))
        config_grid.grid_columnconfigure((0, 1), weight=1)
        
        # Material type
        material_frame = ctk.CTkFrame(config_grid, fg_color="transparent")
        material_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        
        ctk.CTkLabel(
            material_frame,
            text="Tipo de Material:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", pady=(0, 10))
        
        material_options = [
            ("🧬 Biológico", 0),
            ("📦 Carga Geral", 1),
            ("⚠️ Infectante", 2),
            ("💊 Med. e Vacinas", 3)
        ]
        
        for text, value in material_options:
            ctk.CTkRadioButton(
                material_frame,
                text=text,
                value=value,
                variable=self.material_type
            ).pack(anchor="w", pady=2)
        
        # Service type
        service_frame = ctk.CTkFrame(config_grid, fg_color="transparent")
        service_frame.grid(row=0, column=1, padx=(10, 0), sticky="nsew")
        
        ctk.CTkLabel(
            service_frame,
            text="Tipo de Serviço:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", pady=(0, 10))
        
        service_options = [
            ("✈️ Próximo Voo", 0),
            ("📋 Convencional", 1)
        ]
        
        for text, value in service_options:
            ctk.CTkRadioButton(
                service_frame,
                text=text,
                value=value,
                variable=self.flight_service
            ).pack(anchor="w", pady=2)
        
        # Volume and weight
        vol_weight_frame = ctk.CTkFrame(service_frame, fg_color="transparent")
        vol_weight_frame.pack(fill="x", pady=(20, 0))
        
        # Volumes
        vol_container = ctk.CTkFrame(vol_weight_frame, fg_color="transparent")
        vol_container.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(
            vol_container,
            text="Volumes:",
            font=ctk.CTkFont(size=12)
        ).pack(side="left")
        
        self.vols_spinbox = ttk.Spinbox(
            vol_container,
            textvariable=self.vols,
            from_=1,
            to=50,
            width=10
        )
        self.vols_spinbox.pack(side="right")
        
        # Weight
        weight_container = ctk.CTkFrame(vol_weight_frame, fg_color="transparent")
        weight_container.pack(fill="x")
        
        ctk.CTkLabel(
            weight_container,
            text="Peso (kg):",
            font=ctk.CTkFont(size=12)
        ).pack(side="left")
        
        self.weight_spinbox = ttk.Spinbox(
            weight_container,
            textvariable=self.kg_record,
            from_=1,
            to=1000,
            width=10
        )
        self.weight_spinbox.pack(side="right")
        
        # Bulk operations
        bulk_card = ModernCard(self.minutas_frame, title="Operações em Lote")
        bulk_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        bulk_frame = ctk.CTkFrame(bulk_card, fg_color="transparent")
        bulk_frame.pack(fill="x", padx=20, pady=(0, 15))
        bulk_frame.grid_columnconfigure((0, 1), weight=1)
        
        self.emit_bulk_btn = ModernButton(
            bulk_frame,
            text="📄📄 Emitir Minutas Aéreas",
            style="primary",
            command=self.emit_bulk_minutas
        )
        self.emit_bulk_btn.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        
        self.download_folder_btn = ModernButton(
            bulk_frame,
            text="📁 Pasta Downloads",
            style="secondary",
            command=self.select_download_folder
        )
        self.download_folder_btn.grid(row=0, column=1, padx=(10, 0), sticky="ew")
        
        # Progress bar
        self.minutas_progress = ctk.CTkProgressBar(
            self.minutas_frame,
            mode='indeterminate'
        )
        self.minutas_progress.pack(fill="x", padx=20, pady=(10, 20))
    
    def create_cargas_section(self):
        """Create the cargas management section"""
        self.cargas_frame = ModernCard(self.content, title="Gerenciamento de Cargas")
        
        # Actions
        actions_card = ModernCard(self.cargas_frame, title="Ações Disponíveis")
        actions_card.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
        
        actions_container = ctk.CTkFrame(actions_card, fg_color="transparent")
        actions_container.pack(fill="x", padx=20, pady=(0, 15))
        actions_container.grid_columnconfigure(0, weight=1)
        
        self.update_cargas_btn = ModernButton(
            actions_container,
            text="🔄 Atualizar Relação de Cargas",
            style="primary",
            command=self.update_cargas_data
        )
        self.update_cargas_btn.grid(row=0, column=0, pady=(0, 10), sticky="ew")
        
        self.update_fleury_btn = ModernButton(
            actions_container,
            text="📊 Atualizar Planilha Fleury",
            style="secondary",
            command=self.update_fleury_data
        )
        self.update_fleury_btn.grid(row=1, column=0, sticky="ew")
        
        # Progress bar
        self.cargas_progress = ctk.CTkProgressBar(
            self.cargas_frame,
            mode='indeterminate'
        )
        self.cargas_progress.pack(fill="x", padx=20, pady=(10, 20))
    
    def create_arquivos_section(self):
        """Create the file management section"""
        self.arquivos_frame = ModernCard(self.content, title="Gerenciamento de Arquivos")
        
        # Create browse instances
        self.browse1 = Browse(None)
        self.browse2 = Browse(None)
        self.browse3 = Browse(None)
        self.browse5 = Browse(None)
        
        # File selectors
        file_selectors = [
            ("📊 Planilha de Serviços", self.filename, self.browse_services_file),
            ("📦 Relação de Cargas", self.filename2, self.browse_cargas_file),
            ("📁 Pasta das Minutas", self.folderpath, self.browse_minutas_folder),
            ("📋 Planilha Fleury", self.fleury_sheet_name, self.browse_fleury_file)
        ]
        
        for label, var, handler in file_selectors:
            selector = FileSelector(
                self.arquivos_frame,
                label,
                var,
                handler
            )
            selector.pack(fill="x", padx=20, pady=(0, COMPONENT_SPACING))
            
            # Store reference for updates
            if "Serviços" in label:
                self.services_selector = selector
            elif "Cargas" in label:
                self.cargas_selector = selector
            elif "Pasta" in label:
                self.folder_selector = selector
            elif "Fleury" in label:
                self.fleury_selector = selector
    
    def update_nav_selection(self):
        """Update navigation button selection"""
        for i, btn in enumerate(self.nav_buttons):
            if i == ["services", "minutas", "cargas", "arquivos"].index(self.current_section):
                btn.configure(fg_color=COLORS['accent'])
            else:
                btn.configure(fg_color=COLORS['primary'])
    
    def hide_all_sections(self):
        """Hide all content sections"""
        for frame in [self.services_frame, self.minutas_frame, self.cargas_frame, self.arquivos_frame]:
            frame.pack_forget()
    
    def show_services(self):
        """Show services section"""
        self.hide_all_sections()
        self.current_section = "services"
        self.page_title.configure(text="Gerenciamento de Serviços")
        self.services_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.update_nav_selection()
    
    def show_minutas(self):
        """Show minutas section"""
        self.hide_all_sections()
        self.current_section = "minutas"
        self.page_title.configure(text="Emissão de Minutas")
        self.minutas_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.update_nav_selection()
    
    def show_cargas(self):
        """Show cargas section"""
        self.hide_all_sections()
        self.current_section = "cargas"
        self.page_title.configure(text="Gerenciamento de Cargas")
        self.cargas_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.update_nav_selection()
    
    def show_arquivos(self):
        """Show arquivos section"""
        self.hide_all_sections()
        self.current_section = "arquivos"
        self.page_title.configure(text="Gerenciamento de Arquivos")
        self.arquivos_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.update_nav_selection()
    
    # Action methods
    def update_services_data(self):
        """Update services data"""
        self.status_indicators['services_status'].update_status('processing')
        self.thread_0.start_thread(
            services_api,
            self.services_progress,
            arguments=[
                self,
                self.cal_start,
                self.cal_end,
                self.services_frame,
                self.date_filter.get(),
                self.reg_filter.get(),
                self.reg.get(),
                self.filename.get()
            ]
        )
    
    def clear_services_data(self):
        """Clear services data"""
        self.thread_0.start_thread(
            clear_data,
            self.services_progress,
            arguments=[
                self.filename.get(),
                "COLETAS;A2:P500",
                "EMBARQUES;A2:N500",
                "ENTREGAS;A2:L1000",
                "CARGAS;A2:S2000"
            ]
        )
    
    def emit_single_minuta(self):
        """Emit single minuta"""
        self.status_indicators['minutas_status'].update_status('processing')
        self.thread_1.start_thread(
            minutas_api,
            self.minutas_progress,
            arguments=[
                self.dispatch_prot.get(),
                False,
                "",
                self.flight_service.get(),
                self.material_type.get(),
                self.vols.get(),
                self.kg_record.get(),
                self.folderpath.get(),
                self.downloadpath.get(),
                self.dispatch_prot
            ]
        )
    
    def emit_bulk_minutas(self):
        """Emit bulk minutas"""
        self.status_indicators['minutas_status'].update_status('processing')
        self.thread_1.start_thread(
            minutas_all_api,
            self.minutas_progress,
            arguments=[
                self.cal_start,
                self.cal_end,
                self.folderpath.get(),
                self.downloadpath.get(),
                self.dispatch_prot,
                self
            ]
        )
    
    def select_download_folder(self):
        """Select download folder"""
        self.browse3.browse_folder(
            self.downloadpath,
            f'{temp_folder}/filepaths/downloadpath.txt'
        )
    
    def update_cargas_data(self):
        """Update cargas data"""
        self.status_indicators['cargas_status'].update_status('processing')
        self.thread_2.start_thread(
            cargas_api,
            self.cargas_progress,
            arguments=[self.filename2.get()]
        )
    
    def update_fleury_data(self):
        """Update fleury data"""
        self.thread_2.start_thread(
            fleury_sheet,
            self.cargas_progress,
            arguments=[self.cal_start, self.fleury_sheet_name.get()]
        )
    
    def browse_services_file(self):
        """Browse services file"""
        self.browse1.browse_files(
            self.filename,
            f'{temp_folder}/filepaths/sheet_name.txt'
        )
        self.services_selector.update_display(self.filename.get())
    
    def browse_cargas_file(self):
        """Browse cargas file"""
        self.browse2.browse_files(
            self.filename2,
            f'{temp_folder}/filepaths/filename2.txt'
        )
        self.cargas_selector.update_display(self.filename2.get())
    
    def browse_minutas_folder(self):
        """Browse minutas folder"""
        self.browse3.browse_folder(
            self.folderpath,
            f'{temp_folder}/filepaths/folderpath.txt'
        )
        self.folder_selector.update_display(self.folderpath.get())
    
    def browse_fleury_file(self):
        """Browse fleury file"""
        self.browse5.browse_files(
            self.fleury_sheet_name,
            f'{temp_folder}/filepaths/fleury_sheet_name.txt'
        )
        self.fleury_selector.update_display(self.fleury_sheet_name.get())


def main():
    """Main entry point"""
    app = LogLifeModernApp()
    app.mainloop()


if __name__ == "__main__":
    main()
