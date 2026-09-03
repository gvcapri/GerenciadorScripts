from pathlib import Path
from typing import Dict, Any, Callable
from backend.core.base_tool import BaseTool

class ToolAnalisarOrdemConfrontantes(BaseTool):
    name = "Analisar Ordem de Confrontantes"
    description = "Analisa a ordem Frente/Direita/Fundo/Esquerda nos memoriais de lotes e gera um relatório em Excel."
    category = "Análises"
    
    inputs = [
        {"id": "arquivo_docx", "label": "Documento DOCX", "type": "document"}
    ]
    
    def execute(self, inputs: Dict[str, Path], output_dir: Path, progress_callback: Callable[[int, str], None], output_filename: str = "") -> Dict[str, Any]:
        from backend.utils.analisar_ordem_confrontantes import ler_memoriais, salvar_xlsx
        
        caminho_docx = inputs["arquivo_docx"]
        
        progress_callback(30, "Lendo memoriais descritivos no documento...")
        registros = ler_memoriais(caminho_docx)
        
        if not registros:
            raise ValueError("Nenhum memorial encontrado no documento. Verifique o formato.")
            
        progress_callback(70, f"Gerando planilha de análise para {len(registros)} lotes...")
        nome_saida = caminho_docx.stem + "_Analise_Confrontantes.xlsx"
        caminho_xlsx = self.get_output_filepath(output_dir, nome_saida, output_filename)
        
        salvar_xlsx(registros, caminho_xlsx)
        
        progress_callback(100, f"Análise concluída com sucesso! Relatório salvo.")
        return {"status": "success", "output_file": str(caminho_xlsx)}
