from pathlib import Path
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QLabel, QProgressBar, QFrame, QScrollArea, QSizePolicy, QLineEdit, QFileDialog)
from PySide6.QtCore import Qt, QThread, Signal, QSize
from frontend.components.card import Card
from frontend.components.file_selector import FileSelector
from frontend.components.log_console import LogConsole
from frontend.core.theme import theme_manager
from config.settings import settings
import qtawesome as qta
import re

class ToolWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(dict)
    error = Signal(str)

    def __init__(self, tool_instance, inputs, output_dir, output_filename=""):
        super().__init__()
        self.tool = tool_instance
        self.inputs = inputs
        self.output_dir = output_dir
        self.output_filename = output_filename

    def run(self):
        try:
            result = self.tool.execute(self.inputs, self.output_dir, self.progress.emit, self.output_filename)
            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))

class ToolView(QWidget):
    def __init__(self, tool_class, parent=None):
        super().__init__(parent)
        self.tool_class = tool_class
        self.tool_instance = tool_class()
        self.inputs = {}
        self.selectors = {}
        
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(32, 24, 32, 24)
        self.layout.setSpacing(24)
        
        # LEFT COLUMN (65%)
        self.left_col = QWidget()
        left_layout = QVBoxLayout(self.left_col)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(24)
        left_layout.setAlignment(Qt.AlignTop)
        
        # Inputs Card
        if self.tool_instance.inputs:
            inputs_card = Card()
            inputs_header = QHBoxLayout()
            icon_inputs = QLabel()
            icon_inputs.setPixmap(qta.icon('mdi.folder-open-outline', color=theme_manager.get_color('text_main')).pixmap(20, 20))
            self.inputs_title = QLabel("Arquivos e Configurações")
            self.inputs_title.setStyleSheet("font-size: 16px; font-weight: 600;")
            inputs_header.addWidget(icon_inputs)
            inputs_header.addWidget(self.inputs_title)
            inputs_header.addStretch()
            inputs_card.addLayout(inputs_header)
            
            for inp in self.tool_instance.inputs:
                selector = FileSelector(inp.get("label"), inp.get("type", "file"))
                selector.file_selected.connect(lambda path, id=inp["id"]: self.on_file_selected(id, path))
                self.selectors[inp["id"]] = selector
                inputs_card.addWidget(selector)
            
            left_layout.addWidget(inputs_card)
            
        # Output Card
        output_card = Card()
        output_header = QHBoxLayout()
        icon_output = QLabel()
        icon_output.setPixmap(qta.icon('mdi.folder-download-outline', color=theme_manager.get_color('text_main')).pixmap(20, 20))
        self.output_title = QLabel("Saída do Arquivo")
        self.output_title.setStyleSheet("font-size: 16px; font-weight: 600;")
        output_header.addWidget(icon_output)
        output_header.addWidget(self.output_title)
        output_header.addStretch()
        output_card.addLayout(output_header)
        
        self.lbl_dir = QLabel("Pasta de destino")
        self.lbl_dir.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {theme_manager.get_color('text_main')};")
        
        dir_layout = QHBoxLayout()
        self.txt_output_dir = QLineEdit()
        self.txt_output_dir.setReadOnly(True)
        self.txt_output_dir.setPlaceholderText("Pasta atual de geração (padrão)")
        self.txt_output_dir.setText(settings.get("last_output_dir", ""))
        self.txt_output_dir.setFixedHeight(32)
        
        self.btn_select_dir = QPushButton("Selecionar Pasta")
        self.btn_select_dir.setCursor(Qt.PointingHandCursor)
        self.btn_select_dir.setFixedHeight(32)
        self.btn_select_dir.clicked.connect(self.select_output_dir)
        
        dir_layout.addWidget(self.txt_output_dir)
        dir_layout.addWidget(self.btn_select_dir)
        
        self.lbl_name_out = QLabel("Nome do arquivo")
        self.lbl_name_out.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {theme_manager.get_color('text_main')};")
        
        self.txt_output_filename = QLineEdit()
        self.txt_output_filename.setPlaceholderText("Ex.: Comparativo_Julho")
        self.txt_output_filename.setText(settings.get("last_output_filename", ""))
        self.txt_output_filename.setFixedHeight(32)
        
        output_card.addWidget(self.lbl_dir)
        output_card.addLayout(dir_layout)
        output_card.addSpacing(8)
        output_card.addWidget(self.lbl_name_out)
        output_card.addWidget(self.txt_output_filename)
        
        left_layout.addWidget(output_card)
            
        # Execution Card
        exec_card = Card()
        exec_header = QHBoxLayout()
        exec_icon = QLabel()
        exec_icon.setPixmap(qta.icon('mdi.play-circle-outline', color=theme_manager.get_color('text_main')).pixmap(20, 20))
        self.exec_title = QLabel("Execução")
        self.exec_title.setStyleSheet("font-size: 16px; font-weight: 600;")
        
        self.btn_exec = QPushButton("Executar Ferramenta")
        self.btn_exec.setIcon(qta.icon('mdi.play', color='white'))
        self.btn_exec.setCursor(Qt.PointingHandCursor)
        self.btn_exec.clicked.connect(self.execute_tool)
        self.btn_exec.setFixedWidth(180)
        self.btn_exec.setFixedHeight(40)
        
        exec_header.addWidget(exec_icon)
        exec_header.addWidget(self.exec_title)
        exec_header.addStretch()
        exec_header.addWidget(self.btn_exec)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedHeight(6)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.hide()
        
        self.log_console = LogConsole()
        
        exec_card.addLayout(exec_header)
        exec_card.addWidget(self.progress_bar)
        exec_card.addWidget(self.log_console)
        
        left_layout.addWidget(exec_card)
        
        # RIGHT COLUMN (35%)
        self.right_col = QWidget()
        right_layout = QVBoxLayout(self.right_col)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(24)
        right_layout.setAlignment(Qt.AlignTop)
        
        # Info Card
        info_card = Card()
        info_header = QHBoxLayout()
        info_icon = QLabel()
        info_icon.setPixmap(qta.icon('mdi.information-outline', color=theme_manager.get_color('text_main')).pixmap(20, 20))
        self.lbl_name = QLabel(self.tool_instance.name)
        self.lbl_name.setStyleSheet("font-size: 16px; font-weight: 600;")
        
        info_header.addWidget(info_icon)
        info_header.addWidget(self.lbl_name)
        info_header.addStretch()
        
        self.lbl_desc = QLabel(self.tool_instance.description)
        self.lbl_desc.setWordWrap(True)
        self.lbl_desc.setStyleSheet(f"font-size: 13px; color: {theme_manager.get_color('text_sec')}; line-height: 1.4;")
        
        info_card.addLayout(info_header)
        info_card.addWidget(self.lbl_desc)
        right_layout.addWidget(info_card)
        
        # Proportions
        self.layout.addWidget(self.left_col, 65)
        self.layout.addWidget(self.right_col, 35)
        
        self.update_theme()
        
    def on_file_selected(self, id: str, path: Path):
        self.inputs[id] = path
        
    def select_output_dir(self):
        path = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Destino")
        if path:
            self.txt_output_dir.setText(path)
        
    def execute_tool(self):
        missing = [inp["label"] for inp in self.tool_instance.inputs if inp["id"] not in self.inputs]
        if missing:
            self.log_console.append_log(f"Preencha os campos obrigatórios: {', '.join(missing)}", "ERROR")
            return
            
        out_dir = self.txt_output_dir.text().strip()
        out_filename = self.txt_output_filename.text().strip()
        
        if re.search(r'[\\/:*?"<>|]', out_filename):
            self.log_console.append_log("O nome informado contém caracteres inválidos.", "ERROR")
            return
            
        if not out_dir:
            if not self.inputs:
                self.log_console.append_log("Selecione um arquivo de entrada ou escolha uma pasta de saída.", "ERROR")
                return
            out_dir = str(list(self.inputs.values())[0].parent)
            
        if out_dir and not Path(out_dir).exists():
            self.log_console.append_log("A pasta de destino selecionada não existe.", "ERROR")
            return
            
        settings.set("last_output_dir", self.txt_output_dir.text().strip())
        settings.set("last_output_filename", out_filename)
            
        self.btn_exec.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.log_console.clear()
        self.log_console.append_log(f"Iniciando {self.tool_instance.name}...", "INFO")
        
        self.worker = ToolWorker(self.tool_instance, self.inputs, Path(out_dir), out_filename)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
        
    def on_progress(self, percent: int, msg: str):
        self.progress_bar.setValue(percent)
        self.log_console.append_log(msg, "INFO")
        
    def on_finished(self, result: dict):
        self.btn_exec.setEnabled(True)
        self.progress_bar.setValue(100)
        self.log_console.append_log("Execução finalizada com sucesso!", "SUCCESS")
        if "output_file" in result:
            self.log_console.append_log(f"Saída: {result['output_file']}", "INFO")
            
    def on_error(self, err: str):
        self.btn_exec.setEnabled(True)
        self.progress_bar.hide()
        self.log_console.append_log(f"Erro de execução: {err}", "ERROR")

    def update_theme(self):
        self.lbl_name.setStyleSheet(f"font-size: 16px; font-weight: 600; color: {theme_manager.get_color('text_main')};")
        self.lbl_desc.setStyleSheet(f"font-size: 13px; color: {theme_manager.get_color('text_sec')};")
        
        if hasattr(self, 'inputs_title'):
            self.inputs_title.setStyleSheet(f"font-size: 16px; font-weight: 600; color: {theme_manager.get_color('text_main')};")
            
        if hasattr(self, 'output_title'):
            self.output_title.setStyleSheet(f"font-size: 16px; font-weight: 600; color: {theme_manager.get_color('text_main')};")
            self.lbl_dir.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {theme_manager.get_color('text_main')};")
            self.lbl_name_out.setStyleSheet(f"font-weight: 600; font-size: 13px; color: {theme_manager.get_color('text_main')};")
            
            line_edit_style = f"background-color: {theme_manager.get_color('bg_base')}; border: 1px solid {theme_manager.get_color('border')}; border-radius: 4px; padding: 0 8px; color: {theme_manager.get_color('text_main')};"
            self.txt_output_dir.setStyleSheet(line_edit_style)
            self.txt_output_filename.setStyleSheet(line_edit_style)
            
            self.btn_select_dir.setStyleSheet(f"""
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
            """)
            
        self.exec_title.setStyleSheet(f"font-size: 16px; font-weight: 600; color: {theme_manager.get_color('text_main')};")
        
        self.btn_exec.setStyleSheet(f"""
            QPushButton {{
                background-color: {theme_manager.get_color('accent_button')};
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {theme_manager.get_color('accent_hover')};
            }}
            QPushButton:disabled {{
                background-color: {theme_manager.get_color('secondary')};
            }}
        """)
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                background-color: {theme_manager.get_color('border')};
                border: none;
                border-radius: 3px;
            }}
            QProgressBar::chunk {{
                background-color: {theme_manager.get_color('accent')};
                border-radius: 3px;
            }}
        """)
