from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable, Any, Dict, List

class BaseTool(ABC):
    """
    Interface base para todas as ferramentas do sistema.
    As ferramentas devem herdar desta classe para serem registradas automaticamente.
    """
    
    # Nome visível na interface
    name: str = "Ferramenta Desconhecida"
    
    # Descrição da ferramenta
    description: str = "Sem descrição."
    
    # Categoria (ex: Comparativos, Validações, Excel)
    category: str = "Outros"
    
    # Extensão do arquivo final se houver (ex: .xlsx, .pdf)
    output_extension: str = ""
    
    # Definição dos campos de entrada
    # Ex: [{"id": "arq1", "label": "Planilha Excel", "type": "excel"}]
    inputs: List[Dict[str, str]] = []

    @abstractmethod
    def execute(self, inputs: Dict[str, Path], output_dir: Path, progress_callback: Callable[[int, str], None], output_filename: str = "") -> Dict[str, Any]:
        """
        Executa a lógica da ferramenta.
        
        :param inputs: Dicionário onde a chave é o `id` definido em `self.inputs` e o valor é o `Path` do arquivo.
        :param output_dir: O diretório onde o arquivo final deve ser salvo.
        :param progress_callback: Uma função para atualizar a interface. Recebe (porcentagem: int, mensagem: str).
        :param output_filename: Nome personalizado para o arquivo de saída. (Opcional)
        :return: Um dicionário com o resultado da execução. Ex: {"status": "success", "message": "Concluído", "output_file": Path(...)}
        """
        pass

    def get_output_filepath(self, output_dir: Path, default_name: str, output_filename: str = "") -> Path:
        """
        Constrói o caminho final do arquivo, respeitando o diretório e o nome customizado.
        Se output_filename estiver preenchido, usa ele em vez do default_name (mas mantém a extensão).
        """
        if not output_filename:
            return output_dir / default_name
            
        extensao = Path(default_name).suffix
        
        # Previne que o usuário adicione a extensão manualmente e fique duplicada
        if extensao and output_filename.endswith(extensao):
            return output_dir / output_filename
            
        return output_dir / f"{output_filename}{extensao}"
