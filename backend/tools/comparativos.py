import traceback
from pathlib import Path
from typing import Dict, Any, Callable

from backend.core.base_tool import BaseTool
from backend.utils.comparativos_core import (
    load_input, compare_sources, compare_fronts, 
    compare_shp_ecoleta, make_metric_sheet, 
    compare_crf_shp, compare_crf_md, write_result
)

class ComparadorBase(BaseTool):
    category = "Comparativos"
    
    def execute(self, inputs: Dict[str, Path], output_dir: Path, progress_callback: Callable[[int, str], None], output_filename: str = "") -> Dict[str, Any]:
        progress_callback(10, "Lendo arquivos de entrada...")
        
        data = {}
        total = len(inputs)
        for i, (label, path) in enumerate(inputs.items()):
            progress_callback(10 + int((i/total)*30), f"Lendo {label}...")
            data[label] = load_input(label, path)
            
        progress_callback(50, "Comparando dados...")
        sheets = self.run_comparison(data)
        
        progress_callback(80, "Gerando arquivo de saída...")
        
        default_name = "Comparativo_" + self.name.replace(" ", "_").replace("-", "").replace("é", "e").replace("á", "a") + ".xlsx"
        output_file = self.get_output_filepath(output_dir, default_name, output_filename)
        
        for path in inputs.values():
            if path.resolve() == output_file.resolve():
                raise ValueError("O arquivo de saída não pode ser o mesmo que a entrada.")
                
        write_result(output_file, sheets)
        
        counts = {label: len(records) for label, records in data.items()}
        counts["comparados"] = len(sheets.get("Comparacao", []))
        
        progress_callback(100, f"Finalizado! Resultados: {counts}")
        return {"status": "success", "output_file": str(output_file)}
        
    def run_comparison(self, data: Dict[str, list]) -> Dict[str, list]:
        raise NotImplementedError()


class ToolSHPMetrica(ComparadorBase):
    name = "SHP x Métrica"
    description = "Compara Quadra, Lote e Área. A ordem principal segue a planilha Métrica."
    inputs = [
        {"id": "SHP", "label": "Shapefile (Excel)", "type": "excel"},
        {"id": "MT", "label": "Planilha Métrica", "type": "excel"}
    ]
    def run_comparison(self, data):
        return compare_sources(data["SHP"], data["MT"], "SHP", "MT", ("area",), "MT")

class ToolMetricaMD(ComparadorBase):
    name = "Métrica x Memorial de Lotes"
    description = "Compara Quadra, Lote, Área e Perímetro e separa registros faltantes ou sobrando."
    inputs = [
        {"id": "MT", "label": "Planilha Métrica", "type": "excel"},
        {"id": "MD", "label": "Memorial de Lotes", "type": "docx"}
    ]
    def run_comparison(self, data):
        return compare_sources(data["MD"], data["MT"], "MD", "MT", ("area", "perimeter"), "MT")

class ToolFrenteMDSHP(ComparadorBase):
    name = "Frente dos lotes - Memorial x SHP"
    description = "Extrai a frente do Memorial de Lotes e compara com Rua/Logradouro/Frente do SHP."
    inputs = [
        {"id": "MD", "label": "Memorial de Lotes", "type": "docx"},
        {"id": "SHP", "label": "Shapefile (Excel)", "type": "excel"}
    ]
    def run_comparison(self, data):
        return compare_fronts(data["MD"], data["SHP"])

class ToolEcoletaSHP(ComparadorBase):
    name = "Ecoleta x SHP"
    description = "Compara cadastro, beneficiário e coordenadas quando disponíveis nas duas planilhas."
    inputs = [
        {"id": "Ecoleta", "label": "Planilha Ecoleta", "type": "excel"},
        {"id": "SHP", "label": "Shapefile (Excel)", "type": "excel"}
    ]
    def run_comparison(self, data):
        return compare_shp_ecoleta(data["SHP"], data["Ecoleta"])

class ToolPlanilhaMetrica(ComparadorBase):
    name = "Planilha Métrica - Métrica x Memorial"
    description = "Gera uma planilha no formato operacional da Métrica e inclui a conferência com o Memorial."
    inputs = [
        {"id": "MT", "label": "Planilha Métrica", "type": "excel"},
        {"id": "MD", "label": "Memorial de Lotes", "type": "docx"}
    ]
    def run_comparison(self, data):
        return make_metric_sheet(data["MT"], data["MD"])

class ToolCRFSHP(ComparadorBase):
    name = "CRF x SHP"
    description = "Compara a Indicação Numérica do CRF com Quadra, Lote e Área do SHP."
    inputs = [
        {"id": "CRF", "label": "Documento CRF", "type": "document"},
        {"id": "SHP", "label": "Shapefile (Excel)", "type": "excel"}
    ]
    def run_comparison(self, data):
        return compare_crf_shp(data["CRF"], data["SHP"])

class ToolCRFMD(ComparadorBase):
    name = "CRF x Memorial de Lotes"
    description = "Compara Quadra, Lote, Área e Perímetro da Indicação Numérica com o Memorial de Lotes."
    inputs = [
        {"id": "CRF", "label": "Documento CRF", "type": "document"},
        {"id": "MD", "label": "Memorial de Lotes", "type": "docx"}
    ]
    def run_comparison(self, data):
        return compare_crf_md(data["CRF"], data["MD"])
