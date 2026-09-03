from pathlib import Path
from typing import Dict, Any, Callable
from backend.core.base_tool import BaseTool

class ToolOrdenarQuadrasLotesExcel(BaseTool):
    name = "Ordenar Quadras e Lotes (Excel)"
    description = "Ordena naturalmente as quadras e lotes em uma planilha .xlsx, preservando fórmulas e formatações."
    category = "Planilhas"
    
    inputs = [
        {"id": "arquivo_excel", "label": "Planilha Excel (.xlsx)", "type": "excel"}
    ]
    
    def execute(self, inputs: Dict[str, Path], output_dir: Path, progress_callback: Callable[[int, str], None], output_filename: str = "") -> Dict[str, Any]:
        from backend.utils.ordenar_quadras_lotes_excel import processar_arquivo
        
        caminho = str(inputs["arquivo_excel"])
        progress_callback(30, "Lendo e ordenando a planilha...")
        
        nome_saida = inputs["arquivo_excel"].stem + "_ORDENADO.xlsx"
        
        # A função processar_arquivo já tem a lógica de nomeação caso usemos pastas,
        # mas vamos definir a pasta de saída e o nome para termos o caminho correto.
        nome_final = output_filename if output_filename else nome_saida
        if not nome_final.casefold().endswith(".xlsx"):
            nome_final += ".xlsx"
            
        resultado = processar_arquivo(caminho, pasta_saida=str(output_dir), nome_saida=nome_final)
        
        if resultado.status == "erro":
            raise RuntimeError(f"Erro ao processar: {resultado.detalhes}")
        elif resultado.status == "ignorado":
            raise RuntimeError(f"Arquivo ignorado: {resultado.detalhes}")
            
        progress_callback(100, f"Planilha processada com sucesso!")
        return {"status": "success", "output_file": str(resultado.destino)}
