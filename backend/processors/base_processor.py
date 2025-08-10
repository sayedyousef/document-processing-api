#backend/processors/base_processor.py


from abc import ABC, abstractmethod
from pathlib import Path

class BaseProcessor(ABC):
    """Base processor interface"""
    
    @abstractmethod
    async def process(self, file_path: Path, output_dir: Path) -> dict:
        """
        Process a single document
        
        Args:
            file_path: Path to input document
            output_dir: Directory to save output files
            
        Returns:
            dict with keys:
                - filename: Original filename
                - output_filename: Name of created file
                - path: Full path to output file
                - success: Boolean
                - Other processor-specific data
        """
        pass

