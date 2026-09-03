import os
import re
import pandas as pd
from docx import Document
from pathlib import Path
from typing import Dict, Any, Callable

try:
    from docx2pdf import convert as converter_para_pdf
    PDF_DISPONIVEL = True
except ImportError:
    PDF_DISPONIVEL = False

from backend.core.base_tool import BaseTool

class ToolGerarTermos(BaseTool):
    name = "Gerar Termos (Word/PDF)"
    description = "Gera múltiplos Termos a partir de um modelo Word (.docx) e uma planilha Excel com os dados."
    category = "Documentos"
    
    inputs = [
        {"id": "planilha_excel", "label": "Planilha de Dados (.xlsx)", "type": "excel"},
        {"id": "modelo_docx", "label": "Modelo (.docx)", "type": "document"}
    ]
    
    TEXTO_ATUAL = {
        "nome":   "Núcleo Urbano Nossa Senhora Aparecida",
        "cidade": "Santo Antônio do Leverger",
        "data":   "29 de Julho de 2026",
    }
    
    COLUNAS_PLANILHA = {
        "nome":   "Nome do Núcleo/Loteamento",
        "cidade": "Cidade",
        "data":   "Data",
    }
    
    ABA_PLANILHA = "Dados"
    
    def execute(self, inputs: Dict[str, Path], output_dir: Path, progress_callback: Callable[[int, str], None], output_filename: str = "") -> Dict[str, Any]:
        planilha_path = str(inputs["planilha_excel"])
        modelo_path = str(inputs["modelo_docx"])
        
        progress_callback(10, "Lendo registros da planilha...")
        registros = self._ler_registros(planilha_path)
        
        if not registros:
            raise RuntimeError(f"Nenhum registro encontrado na aba '{self.ABA_PLANILHA}'.")
            
        total = len(registros)
        
        # Cria uma subpasta para os termos para não misturar com outros outputs
        saida_termos_dir = output_dir / (output_filename if output_filename else "Termos_Gerados")
        saida_termos_dir.mkdir(parents=True, exist_ok=True)
        
        for i, registro in enumerate(registros):
            progress_callback(10 + int((i / total) * 60), f"Gerando documento {i+1} de {total}...")
            
            mapa_substituicao = {
                self.TEXTO_ATUAL["nome"]: registro["nome"],
                self.TEXTO_ATUAL["cidade"]: registro["cidade"],
                self.TEXTO_ATUAL["data"]: registro["data"],
            }
            
            nome_arquivo = f"Termo_{self._nome_arquivo_seguro(registro['nome'])}.docx"
            caminho_saida = saida_termos_dir / nome_arquivo
            
            self._gerar_documento(modelo_path, mapa_substituicao, caminho_saida)
            
        progress_callback(75, "Documentos Word gerados. Convertendo para PDF (pode demorar)...")
        
        if PDF_DISPONIVEL:
            try:
                converter_para_pdf(str(saida_termos_dir))
                progress_callback(95, "Conversão para PDF concluída.")
            except Exception as erro:
                progress_callback(95, f"PDFs não puderam ser gerados: {erro}")
        else:
            progress_callback(95, "Aviso: docx2pdf não instalado, apenas .docx gerados.")
            
        progress_callback(100, f"Gerados {total} documentos na pasta {saida_termos_dir.name}.")
        return {"status": "success", "output_file": str(saida_termos_dir)}

    def _ler_registros(self, planilha_path: str):
        df = pd.read_excel(planilha_path, sheet_name=self.ABA_PLANILHA)
        df = df.dropna(how="all")
        
        registros = []
        for _, linha in df.iterrows():
            registro = {
                chave: str(linha[coluna]).strip()
                for chave, coluna in self.COLUNAS_PLANILHA.items()
            }
            if all(registro.values()) and "nan" not in registro.values():
                registros.append(registro)
        return registros
        
    def _substituir_no_paragrafo(self, paragrafo, mapa):
        runs = paragrafo.runs
        if not runs:
            return
            
        texto_completo = "".join(r.text for r in runs)
        if not any(antigo in texto_completo for antigo in mapa):
            return
            
        limites = []
        pos = 0
        for r in runs:
            limites.append((pos, pos + len(r.text)))
            pos += len(r.text)
            
        ocorrencias = []
        for antigo, novo in mapa.items():
            start = 0
            while True:
                idx = texto_completo.find(antigo, start)
                if idx == -1:
                    break
                ocorrencias.append((idx, idx + len(antigo), novo))
                start = idx + len(antigo)
                
        ocorrencias.sort(key=lambda o: o[0], reverse=True)
        
        for inicio, fim, novo in ocorrencias:
            run_ini = run_fim = None
            for i, (rs, re_) in enumerate(limites):
                if rs < fim and re_ > inicio:
                    if run_ini is None:
                        run_ini = i
                    run_fim = i
            if run_ini is None:
                continue
                
            rs_ini, _ = limites[run_ini]
            rs_fim, _ = limites[run_fim]
            prefixo = runs[run_ini].text[: max(0, inicio - rs_ini)]
            sufixo = runs[run_fim].text[fim - rs_fim:]
            
            if run_ini == run_fim:
                runs[run_ini].text = prefixo + novo + sufixo
            else:
                runs[run_ini].text = prefixo + novo
                runs[run_fim].text = sufixo
                for i in range(run_ini + 1, run_fim):
                    runs[i].text = ""

    def _percorrer_paragrafos(self, container, mapa):
        for paragrafo in container.paragraphs:
            self._substituir_no_paragrafo(paragrafo, mapa)
        for tabela in getattr(container, "tables", []):
            for linha in tabela.rows:
                for celula in linha.cells:
                    self._percorrer_paragrafos(celula, mapa)

    def _gerar_documento(self, modelo_path, mapa, caminho_saida):
        doc = Document(modelo_path)
        self._percorrer_paragrafos(doc, mapa)
        
        for secao in doc.sections:
            self._percorrer_paragrafos(secao.header, mapa)
            self._percorrer_paragrafos(secao.footer, mapa)
            
        doc.save(caminho_saida)
        
    def _nome_arquivo_seguro(self, texto):
        texto = re.sub(r'[\\/:*?"<>|]', "", texto)
        return texto.strip()
