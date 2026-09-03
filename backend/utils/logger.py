import logging
import sys
from pathlib import Path
from datetime import datetime

from config.settings import settings

def get_data_dir():
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent.parent.parent

LOGS_DIR = get_data_dir() / "logs"

def setup_logger():
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    
    current_date = datetime.now().strftime("%Y-%m-%d")
    log_file = LOGS_DIR / f"app_{current_date}.log"
    
    log_level_str = settings.get("log_level", "INFO").upper()
    log_level = getattr(logging, log_level_str, logging.INFO)
    
    # Formatação padronizada
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Handler para arquivo
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    file_handler.setLevel(log_level)
    
    # Handler para console (sys.stdout)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(log_level)
    
    # Configurar logger root
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)
    
    # Evitar adicionar handlers múltiplos em re-execuções
    if not root_logger.handlers:
        root_logger.addHandler(file_handler)
        root_logger.addHandler(console_handler)

    logging.info("Sistema de logs inicializado.")
