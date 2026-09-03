import importlib
import logging
from typing import Dict, Type, List

from backend.core.base_tool import BaseTool

logger = logging.getLogger(__name__)

class ToolRegistry:
    def __init__(self):
        self._tools: Dict[str, Type[BaseTool]] = {}

    def get_all_subclasses(self, cls):
        subclasses = []
        for subclass in cls.__subclasses__():
            subclasses.append(subclass)
            subclasses.extend(self.get_all_subclasses(subclass))
        return subclasses

    def discover_tools(self):
        """
        Descobre ferramentas utilizando __subclasses__() que é 100% seguro com PyInstaller.
        Requer que os módulos tenham sido importados.
        """
        try:
            import backend.tools.comparativos
            import backend.tools.pdf_tools
            import backend.tools.formatacao
            import backend.tools.edicoes
            import backend.tools.planilhas
            import backend.tools.analises
            import backend.tools.termos
        except Exception as e:
            logger.error(f"Erro ao importar ferramentas: {e}")

        for cls in self.get_all_subclasses(BaseTool):
            if cls.name != "Ferramenta Desconhecida":
                self.register_tool(cls)

    def register_tool(self, tool_class: Type[BaseTool]):
        if tool_class.name in self._tools:
            return
        self._tools[tool_class.name] = tool_class
        logger.info(f"Ferramenta registrada: {tool_class.name} ({tool_class.category})")

    def get_tool(self, name: str) -> Type[BaseTool]:
        if name not in self._tools:
            raise KeyError(f"Ferramenta '{name}' não encontrada.")
        return self._tools[name]

    def get_all_tools(self) -> List[Type[BaseTool]]:
        return list(self._tools.values())
        
    def get_categories(self) -> List[str]:
        categories = set(tool.category for tool in self._tools.values())
        return sorted(list(categories))

# Instância global do registry
registry = ToolRegistry()
