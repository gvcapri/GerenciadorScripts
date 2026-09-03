import os
import glob
from pathlib import Path
from typing import Dict, Any, Callable
import fitz  # PyMuPDF

from backend.core.base_tool import BaseTool

class ToolAssinarPDF(BaseTool):
    name = "Assinatura e Logo em Lote (PDF)"
    description = "Processa uma pasta inteira de PDFs inserindo Assinatura e Logo baseado em âncoras de texto."
    category = "PDF"
    
    inputs = [
        {"id": "pasta_pdfs", "label": "Pasta com PDFs", "type": "folder"},
        {"id": "assinatura", "label": "Imagem Assinatura", "type": "image"},
        {"id": "logo", "label": "Imagem Logo", "type": "image"}
    ]
    
    def execute(self, inputs: Dict[str, Path], output_dir: Path, progress_callback: Callable[[int, str], None], output_filename: str = "") -> Dict[str, Any]:
        pasta_pdfs = inputs["pasta_pdfs"]
        caminho_assinatura = str(inputs["assinatura"])
        caminho_logo = str(inputs["logo"])
        
        arquivos_pdf = glob.glob(os.path.join(str(pasta_pdfs), "*.pdf"))
        if not arquivos_pdf:
            raise ValueError(f"Nenhum arquivo PDF encontrado na pasta: {pasta_pdfs}")
            
        total = len(arquivos_pdf)
        sucesso = 0
        erros = 0
        
        for i, arquivo in enumerate(arquivos_pdf):
            nome_arquivo = os.path.basename(arquivo)
            progress_callback(int((i / total) * 100), f"Processando {nome_arquivo}...")
            
            try:
                doc = fitz.open(arquivo)
                pagine = doc[0]
                
                rect_assinatura = None
                rect_logo = None
                
                textos_nome = pagine.search_for("Thiago Costa Marques Ninomiya")
                if textos_nome:
                    r_nome = textos_nome[0] 
                    rect_assinatura = fitz.Rect(r_nome.x0 + 15, r_nome.y0 - 35, r_nome.x1 - 15, r_nome.y0 - 8)
                
                textos_datum = pagine.search_for("DATUM/FUSO/MC")
                if textos_datum:
                    r_datum = textos_datum[0]
                    rect_logo = fitz.Rect(r_datum.x1 + 60, r_datum.y0 - 80, r_datum.x1 + 240, r_datum.y1 + 10)
                else:
                    rect_logo = fitz.Rect(640, 450, 810, 550)

                tem_assinatura = False
                tem_logo = False
                imagens_na_pagina = pagine.get_image_info()
                
                for img in imagens_na_pagina:
                    if "bbox" in img:
                        bbox_imagem = fitz.Rect(img["bbox"])
                        if rect_assinatura and bbox_imagem.intersects(rect_assinatura):
                            tem_assinatura = True
                        if rect_logo and bbox_imagem.intersects(rect_logo):
                            tem_logo = True

                modificado = False
                if rect_assinatura and not tem_assinatura:
                    pagine.insert_image(rect_assinatura, filename=caminho_assinatura)
                    modificado = True

                if rect_logo and not tem_logo:
                    pagine.insert_image(rect_logo, filename=caminho_logo)
                    modificado = True

                if modificado:
                    # Salva na pasta de saída para não corromper o original em caso de erro
                    out_file = output_dir / nome_arquivo
                    doc.save(str(out_file))
                    sucesso += 1
                
                doc.close()

            except Exception as e:
                progress_callback(int((i / total) * 100), f"[ERRO] Falha ao processar {nome_arquivo}: {e}")
                erros += 1
                
        progress_callback(100, f"Processamento finalizado. Sucesso: {sucesso} | Erros: {erros}")
        return {"status": "success", "message": f"Concluído. Salvos em {output_dir}"}
