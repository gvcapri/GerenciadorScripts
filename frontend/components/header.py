from PySide6.QtWidgets import QFrame, QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QLineEdit, QWidget
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QCursor, QAction
from frontend.core.theme import theme_manager
import qtawesome as qta

class Header(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("header")
        self.setFixedHeight(80)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(32, 16, 32, 16)
        
        # Titulo e descricao
        title_layout = QVBoxLayout()
        title_layout.setSpacing(4)
        title_layout.setAlignment(Qt.AlignVCenter)
        self.lbl_title = QLabel("Bem-vindo ao Gerenciador de Scripts")
        self.lbl_title.setStyleSheet("font-size: 24px; font-weight: 600;")
        self.lbl_desc = QLabel("Selecione uma ferramenta no menu lateral para começar.")
        self.lbl_desc.setStyleSheet("font-size: 14px;")
        title_layout.addWidget(self.lbl_title)
        title_layout.addWidget(self.lbl_desc)
        
        # Action Buttons
        self.btn_theme = self._create_icon_btn('mdi.theme-light-dark')
        self.btn_settings = self._create_icon_btn('mdi.cog-outline')
        self.btn_help = self._create_icon_btn('mdi.help-circle-outline')
        self.btn_user = self._create_icon_btn('mdi.account-circle-outline')
        
        layout.addLayout(title_layout)
        layout.addStretch()
        layout.addWidget(self.btn_theme)
        layout.addWidget(self.btn_settings)
        layout.addWidget(self.btn_help)
        layout.addSpacing(16)
        layout.addWidget(self.btn_user)
        
        self.update_theme()
        
    def _create_icon_btn(self, icon_name):
        btn = QPushButton()
        btn.setCursor(QCursor(Qt.PointingHandCursor))
        btn.setFixedSize(36, 36)
        btn.setProperty("icon_name", icon_name)
        return btn
        
    def set_title(self, title: str, desc: str):
        self.lbl_title.setText(title)
        self.lbl_desc.setText(desc)
        
    def update_theme(self):
        self.setStyleSheet(f"""
            #header {{
                background-color: {theme_manager.get_color('bg_card')};
                border-bottom: 1px solid {theme_manager.get_color('border')};
            }}
        """)
        self.lbl_title.setStyleSheet(f"font-size: 24px; font-weight: 600; color: {theme_manager.get_color('text_main')};")
        self.lbl_desc.setStyleSheet(f"font-size: 14px; color: {theme_manager.get_color('text_sec')};")
        
        # Update Icon buttons
        for btn in [self.btn_theme, self.btn_settings, self.btn_help, self.btn_user]:
            icon_name = btn.property("icon_name")
            btn.setIcon(qta.icon(icon_name, color=theme_manager.get_color('text_main')))
            btn.setIconSize(QSize(20, 20))
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    border: none;
                    border-radius: 18px;
                }}
                QPushButton:hover {{
                    background-color: {theme_manager.get_color('primary_hover')};
                }}
            """)
