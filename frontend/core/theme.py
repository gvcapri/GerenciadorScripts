from enum import Enum
from typing import Dict
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QApplication

class ThemeMode(Enum):
    LIGHT = "light"
    DARK = "dark"

PALETTES = {
    ThemeMode.LIGHT: {
        "primary": "#233A5E",
        "primary_hover": "#334155",
        "secondary": "#64748B",
        "bg_base": "#FFFFFF",
        "bg_card": "#F8FAFC",
        "border": "#E2E8F0",
        "accent": "#91B522",
        "accent_button": "#7F9F1F",
        "accent_hover": "#6F8D19",
        "text_main": "#111827",
        "text_sec": "#4B5563",
        "sidebar_bg": "#F3F4F6",
        "sidebar_text": "#1F2937",
        "sidebar_text_sec": "#6B7280",
        "sidebar_hover": "#E5E7EB",
        "sidebar_active_bg": "#FFFFFF",
        "sidebar_active_text": "#7F9F1F"
    },
    ThemeMode.DARK: {
        "primary": "#F2F4F7",
        "primary_hover": "#2C3440",
        "secondary": "#B9C2CF",
        "bg_base": "#1A1E24",
        "bg_card": "#232830",
        "border": "#40516F",
        "accent": "#91B522",
        "accent_button": "#7F9F1F",
        "accent_hover": "#6F8D19",
        "text_main": "#E1E5EB",
        "text_sec": "#9EA8B6",
        "sidebar_bg": "#15191E",
        "sidebar_text": "#F2F4F7",
        "sidebar_text_sec": "#9EA8B6",
        "sidebar_hover": "#2C3440",
        "sidebar_active_bg": "#0D1013",
        "sidebar_active_text": "#91B522"
    }
}

class ThemeManager:
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance.current_mode = ThemeMode.LIGHT
        return cls._instance

    def get_color(self, name: str) -> str:
        return PALETTES[self.current_mode][name]
        
    def toggle_theme(self):
        if self.current_mode == ThemeMode.DARK:
            self.current_mode = ThemeMode.LIGHT
        else:
            self.current_mode = ThemeMode.DARK
        self.apply_theme()
        
    def apply_theme(self):
        app = QApplication.instance()
        if not app:
            return
            
        qss = f"""
        QWidget {{
            font-family: 'Inter', 'Segoe UI', sans-serif;
            color: {self.get_color('text_main')};
        }}
        QMainWindow, #central_widget, #right_container, QStackedWidget {{
            background-color: {self.get_color('bg_base')};
        }}
        QScrollArea, QScrollArea > QWidget > QWidget {{
            background-color: transparent;
            border: none;
        }}
        QScrollBar:vertical {{
            border: none;
            background: transparent;
            width: 6px;
            margin: 0px;
        }}
        QScrollBar::handle:vertical {{
            background: {self.get_color('border')};
            min-height: 30px;
            border-radius: 3px;
        }}
        QScrollBar::handle:vertical:hover {{
            background: {self.get_color('text_sec')};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            border: none;
            background: none;
        }}
        """
        app.setStyleSheet(qss)
        
theme_manager = ThemeManager()
