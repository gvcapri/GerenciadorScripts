from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel
from PySide6.QtCore import Qt
from frontend.core.theme import theme_manager

class StatusBar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("status_bar")
        self.setFixedHeight(32)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 0, 16, 0)
        
        self.lbl_version = QLabel("v2.0.0")
        self.lbl_mode = QLabel(theme_manager.current_mode.value.upper())
        self.lbl_status = QLabel("Pronto")
        self.lbl_modules = QLabel("0 Módulos")
        
        for lbl in [self.lbl_version, self.lbl_mode, self.lbl_modules, self.lbl_status]:
            lbl.setStyleSheet("font-size: 12px;")
            
        layout.addWidget(self.lbl_version)
        layout.addWidget(self.create_separator())
        layout.addWidget(self.lbl_mode)
        layout.addWidget(self.create_separator())
        layout.addWidget(self.lbl_modules)
        layout.addStretch()
        layout.addWidget(self.lbl_status)
        
        self.update_theme()
        
    def create_separator(self):
        lbl = QLabel("•")
        lbl.setStyleSheet(f"color: {theme_manager.get_color('border')}; margin: 0 8px;")
        return lbl
        
    def update_theme(self):
        self.setStyleSheet(f"""
            #status_bar {{
                background-color: {theme_manager.get_color('bg_card')};
                color: {theme_manager.get_color('text_sec')};
                border-top: 1px solid {theme_manager.get_color('border')};
            }}
            QLabel {{
                color: {theme_manager.get_color('text_sec')};
            }}
        """)
        self.lbl_mode.setText(theme_manager.current_mode.value.upper())
        
    def set_modules_count(self, count: int):
        self.lbl_modules.setText(f"{count} Módulos Carregados")
