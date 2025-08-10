# backend/processors/full-word-processor/WordFullProcessor.py

from pathlib import Path
import sys

# Add parent path for base processor
sys.path.append(str(Path(__file__).parent.parent))

from processors.base_processor import BaseProcessor
from .document_extractor import DocumentExtractor
from .footnote_handler import FootnoteHandler
from .image_extractor import ImageExtractor
from .mammoth_handler import MammothHandler
from .html_generator import HTMLGenerator
from .models import ProcessingContext

class WordFullProcessor(BaseProcessor):
    """
    Word to HTML processor for documents with equations already processed.
    Handles: document parsing, footnotes, images, and HTML generation
    """
    
    def __init__(self):
        super().__init__()
        self.document_extractor = DocumentExtractor()
        self.footnote_handler = FootnoteHandler()
        self.image_extractor = ImageExtractor()
        self.mammoth_handler = MammothHandler()
        self.html_generator = HTMLGenerator()
        
    def process_document(self, input_file, output_dir):
        """
        Process Word document to HTML
        NOTE: Expects document with equations already processed
        """
        # Initialize context
        context = ProcessingContext(
            input_path=Path(input_file),
            output_dir=Path(output_dir)
        )
        
        # The working document is the input (equations already processed)
        context.working_doc = context.input_path
        
        print(f"\n{'='*60}")
        print(f"WORD TO HTML PROCESSING")
        print(f"Input: {context.input_path.name}")
        print(f"Output dir: {context.output_dir}")
        print(f"{'='*60}\n")
        
        # Step 1: Extract document structure
        context = self.document_extractor.extract(context)
        
        # Step 2: Process footnotes
        context = self.footnote_handler.process(context)
        
        # Step 3: Setup image extraction
        context = self.image_extractor.setup(context)
        
        # Step 4: Convert with mammoth
        context = self.mammoth_handler.convert(context)
        
        # Step 5: Generate final HTML
        output_file = self.html_generator.generate(context)
        
        print(f"\n{'='*60}")
        print(f"✅ HTML PROCESSING COMPLETE")
        print(f"📄 Output: {output_file}")
        print(f"📁 Images: {context.images_dir}")
        print(f"📝 Footnotes: {len(context.footnotes)}")
        print(f"{'='*60}\n")
        
        return output_file