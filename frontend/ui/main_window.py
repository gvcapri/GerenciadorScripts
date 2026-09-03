import sys
from pathlib import Path
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QStackedWidget, QScrollArea, QFrame
from frontend.core.theme import theme_manager
from frontend.components.sidebar import Sidebar
from frontend.components.header import Header
from frontend.components.status_bar import StatusBar
from frontend.views.tool_view import ToolView
from backend.core.registry import registry

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciador de Scripts - GEOGIS")
        self.setMinimumSize(1100, 750)
        
        # Central widget and main layout
        central_widget = QWidget()
        central_widget.setObjectName("central_widget")
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Sidebar
        self.sidebar = Sidebar()
        main_layout.addWidget(self.sidebar)
        
        # Right container
        right_container = QWidget()
        right_container.setObjectName("right_container")
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)
        main_layout.addWidget(right_container)
        
        # Header
        self.header = Header()
        self.header.btn_theme.clicked.connect(self.toggle_theme)
        right_layout.addWidget(self.header)
        
        # Content Area (Stacked Widget inside ScrollArea)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setObjectName("main_scroll")
        
        self.stacked_widget = QStackedWidget()
        self.scroll_area.setWidget(self.stacked_widget)
        right_layout.addWidget(self.scroll_area)
        
        # Status bar
        self.status_bar = StatusBar()
        right_layout.addWidget(self.status_bar)
        
        self.populate_sidebar()
        theme_manager.apply_theme()
        
    def populate_sidebar(self):
        categories = registry.get_categories()
        tools = registry.get_all_tools()
        
        self.views = []
        
        for category in categories:
            self.sidebar.add_category(category)
            cat_tools = [t for t in tools if t.category == category]
            
            for tool_class in cat_tools:
                view = ToolView(tool_class)
                index = self.stacked_widget.addWidget(view)
                self.views.append(view)
                self.sidebar.add_tool(tool_class.name, tool_class, index)
                
        self.sidebar.item_clicked.connect(self.on_tool_selected)
        self.status_bar.set_modules_count(len(tools))
        
    def on_tool_selected(self, index: int, tool_class):
        self.stacked_widget.setCurrentIndex(index)
        self.header.set_title(tool_class.name, tool_class.description)
        
    def toggle_theme(self):
        theme_manager.toggle_theme()
        self.header.update_theme()
        self.sidebar.update_theme()
        self.status_bar.update_theme()
        for view in self.views:
            view.update_theme()
            for child in view.findChildren(QWidget):
                if hasattr(child, "update_theme"):
                    child.update_theme()
