from PySide6.QtWidgets import QFrame, QVBoxLayout, QGraphicsDropShadowEffect
from PySide6.QtGui import QColor
from frontend.core.theme import theme_manager

class Card(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(24, 24, 24, 24)
        self.layout.setSpacing(16)
        
        self.shadow = QGraphicsDropShadowEffect(self)
        self.shadow.setBlurRadius(20)
        self.shadow.setXOffset(0)
        self.shadow.setYOffset(2)
        self.setGraphicsEffect(self.shadow)
        
        self.update_theme()
        
    def addWidget(self, widget):
        self.layout.addWidget(widget)
        
    def addLayout(self, layout):
        self.layout.addLayout(layout)
        
    def addSpacing(self, spacing: int):
        self.layout.addSpacing(spacing)
        
    def update_theme(self):
        self.setStyleSheet(f"""
            #card {{
                background-color: {theme_manager.get_color('bg_card')};
                border: 1px solid {theme_manager.get_color('border')};
                border-radius: 8px;
            }}
        """)
        shadow_color = QColor(0, 0, 0, 15) if theme_manager.current_mode.value == "light" else QColor(0, 0, 0, 40)
        self.shadow.setColor(shadow_color)
