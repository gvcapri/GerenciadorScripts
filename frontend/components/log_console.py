from PySide6.QtWidgets import QWidget, QVBoxLayout, QTextEdit, QHBoxLayout, QPushButton, QLabel, QFrame
from PySide6.QtGui import QFont, QTextCursor
from PySide6.QtCore import Qt, QSize
from frontend.core.theme import theme_manager
import qtawesome as qta

class LogConsole(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("log_console")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)
        
        # Header
        self.header = QWidget()
        header_layout = QHBoxLayout(self.header)
        header_layout.setContentsMargins(12, 6, 12, 6)
        
        self.lbl_title = QLabel("TERMINAL")
        
        self.btn_clear = QPushButton()
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.setToolTip("Limpar Terminal")
        self.btn_clear.setFixedSize(24, 24)
        self.btn_clear.clicked.connect(self.clear)
        
        header_layout.addWidget(self.lbl_title)
        header_layout.addStretch()
        header_layout.addWidget(self.btn_clear)
        
        # Text Edit
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setFont(QFont("Consolas", 10))
        self.text_edit.setMinimumHeight(150)
        self.text_edit.setMaximumHeight(250)
        
        self.layout.addWidget(self.header)
        self.layout.addWidget(self.text_edit)
        self.update_theme()
        
    def update_theme(self):
        self.setStyleSheet(f"""
            #log_console {{
                border: 1px solid {theme_manager.get_color('border')};
                border-radius: 6px;
                background-color: {theme_manager.get_color('bg_base')};
            }}
        """)
        
        self.header.setStyleSheet(f"""
            border-bottom: 1px solid {theme_manager.get_color('border')};
        """)
        
        self.lbl_title.setStyleSheet(f"""
            font-weight: 600; 
            font-size: 11px; 
            color: {theme_manager.get_color('text_sec')};
            letter-spacing: 1px;
            border: none;
        """)
        
        self.btn_clear.setIcon(qta.icon('mdi.trash-can-outline', color=theme_manager.get_color('text_sec')))
        self.btn_clear.setIconSize(QSize(16, 16))
        self.btn_clear.setStyleSheet(f"""
            QPushButton {{
                background: transparent;
                border: none;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background: {theme_manager.get_color('primary_hover')};
            }}
        """)
        
        self.text_edit.setStyleSheet(f"""
            QTextEdit {{
                background-color: transparent;
                color: {theme_manager.get_color('text_main')};
                border: none;
                padding: 8px;
                selection-background-color: {theme_manager.get_color('primary_hover')};
            }}
        """)
        
    def append_log(self, text: str, level: str = "INFO"):
        color_map = {
            "INFO": theme_manager.get_color('text_sec'),
            "SUCCESS": theme_manager.get_color('accent'),
            "WARNING": "#Eab308",
            "ERROR": "#Ef4444"
        }
        color = color_map.get(level.upper(), color_map["INFO"])
        html = f'<span style="color: {color};">[{level}]</span> <span style="color: {theme_manager.get_color("text_main")};">{text}</span><br>'
        
        self.text_edit.moveCursor(QTextCursor.End)
        self.text_edit.insertHtml(html)
        self.text_edit.ensureCursorVisible()
        
    def clear(self):
        self.text_edit.clear()
