# backend/processors/full-word-processor/models.py

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Dict, Optional

@dataclass
class Footnote:
    """Footnote data model"""
    id: str
    reference_id: str
    content: str
    anchor: str = ""
    
@dataclass
class ImageInfo:
    """Image information model"""
    filename: str
    path: Path
    content_type: str
    
@dataclass
class ProcessingContext:
    """Context passed through processing pipeline"""
    input_path: Path
    output_dir: Path
    working_doc: Optional[Path] = None
    images_dir: Optional[Path] = None
    footnotes: List[Footnote] = field(default_factory=list)
    images: List[ImageInfo] = field(default_factory=list)
    html_content: str = ""
    metadata: Dict = field(default_factory=dict)