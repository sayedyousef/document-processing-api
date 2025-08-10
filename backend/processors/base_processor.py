from abc import ABC, abstractmethod
from pathlib import Path

class BaseProcessor(ABC):
    """Base processor interface"""
    
    @abstractmethod
    async def process(self, file_path: Path) -> dict:
        """Process a single document"""
        pass