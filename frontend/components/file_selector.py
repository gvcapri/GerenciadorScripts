import os
from pathlib import Path
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFileDialog, QFrame, QSizePolicy)
from PySide6.QtCore import Qt, Signal, QSize
from frontend.core.theme import theme_manager
import qtawesome as qta

class FileSelector(QWidget):
    file_selected = Signal(Path)
    
    def __init__(self, label: str, input_type: str = "file", parent=None):
        super().__init__(parent)
        self.input_type = input_type
        self.current_path = None
        self.label_text = label
        
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(6)
        
        self.lbl_title = QLabel(label)
        self.layout.addWidget(self.lbl_title)
        
        self.container = QFrame()
        self.container.setObjectName("file_container")
        self.container_layout = QHBoxLayout(self.container)
        self.container_layout.setContentsMargins(12, 8, 12, 8)
        self.container_layout.setSpacing(12)
        
        self.lbl_icon = QLabel()
        self.lbl_icon.setFixedSize(24, 24)
        
        self.info_layout = QVBoxLayout()
        self.info_layout.setSpacing(2)
        
        self.lbl_name = QLabel("Nenhum arquivo selecionado")
        self.lbl_path = QLabel("Clique em Selecionar para começar")
        self.lbl_path.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.info_layout.addWidget(self.lbl_name)
        self.info_layout.addWidget(self.lbl_path)
        
        self.btn_select = QPushButton("Selecionar")
        self.btn_select.setCursor(Qt.PointingHandCursor)
        self.btn_select.clicked.connect(self.open_dialog)
        self.btn_select.setFixedHeight(32)
        
        self.btn_clear = QPushButton()
        self.btn_clear.setCursor(Qt.PointingHandCursor)
        self.btn_clear.setToolTip("Remover arquivo")
        self.btn_clear.clicked.connect(self.clear_selection)
        self.btn_clear.setFixedSize(32, 32)
        self.btn_clear.hide()
        
        self.container_layout.addWidget(self.lbl_icon)
        self.container_layout.addLayout(self.info_layout, 1)
        self.container_layout.addWidget(self.btn_select)
        self.container_layout.addWidget(self.btn_clear)
        
        self.layout.addWidget(self.container)
        self.update_theme()
        
    def update_theme(self):
        self.lbl_title.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {theme_manager.get_color('text_main')};")
        
        is_empty = self.current_path is None
        border_style = f"1px dashed {theme_manager.get_color('border')}" if is_empty else f"1px solid {theme_manager.get_color('border')}"
        bg_color = "transparent" if is_empty else theme_manager.get_color('bg_base')
        
        self.container.setStyleSheet(f"""
            #file_container {{
                background-color: {bg_color};
                border: {border_style};
                border-radius: 6px;
            }}
        """)
        
        icon_name = 'mdi.folder-outline' if self.input_type == 'folder' else 'mdi.file-document-outline'
        if not is_empty:
            icon_name = 'mdi.folder' if self.input_type == 'folder' else 'mdi.file-document'
            
        icon_color = theme_manager.get_color('text_sec') if is_empty else theme_manager.get_color('accent')
        self.lbl_icon.setPixmap(qta.icon(icon_name, color=icon_color).pixmap(24, 24))
        
        self.lbl_name.setStyleSheet(f"font-weight: 500; font-size: 13px; color: {theme_manager.get_color('text_main') if not is_empty else theme_manager.get_color('text_sec')};")
        self.lbl_path.setStyleSheet(f"font-size: 11px; color: {theme_manager.get_color('text_sec')};")
        
        btn_style = f"""
            QPushButton {{
                background-color: transparent;
                border: 1px solid {theme_manager.get_color('border')};
                border-radius: 4px;
                padding: 0 12px;
                font-weight: 500;
                color: {theme_manager.get_color('text_main')};
            }}
            QPushButton:hover {{
                background-color: {theme_manager.get_color('primary_hover')};
            }}
        """
        self.btn_select.setStyleSheet(btn_style)
        
        self.btn_clear.setIcon(qta.icon('mdi.close', color=theme_manager.get_color('text_sec')))
        self.btn_clear.setIconSize(QSize(16, 16))
        self.btn_clear.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                border: none;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                background-color: {theme_manager.get_color('primary_hover')};
            }}
        """)
        
    def open_dialog(self):
        if self.input_type == "folder":
            path = QFileDialog.getExistingDirectory(self, f"Selecione a pasta para {self.label_text}")
        else:
            path, _ = QFileDialog.getOpenFileName(self, f"Selecione {self.label_text}")
            
        if path:
            self.set_path(path)
            
    def set_path(self, path: str):
        self.current_path = Path(path)
        self.lbl_name.setText(self.current_path.name)
        self.lbl_path.setText(str(self.current_path))
        self.btn_select.setText("Alterar")
        self.btn_clear.show()
        self.update_theme()
        self.file_selected.emit(self.current_path)
        
    def clear_selection(self):
        self.current_path = None
        self.lbl_name.setText("Nenhum arquivo selecionado")
        self.lbl_path.setText("Clique em Selecionar para começar")
        self.btn_select.setText("Selecionar")
        self.btn_clear.hide()
        self.update_theme()
