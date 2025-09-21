# backend/processors/full-word-processor/mammoth_handler.py

import mammoth
from .models import ProcessingContext

class MammothHandler:
    """Handles mammoth conversion with footnotes and images"""
    
    def __init__(self):
        self.image_extractor = None
        
    def convert(self, context: ProcessingContext) -> ProcessingContext:
        """Convert document to HTML using mammoth"""
        
        print("\nSTEP 5: Converting to HTML with mammoth...")
        
        # Import image extractor
        from .image_extractor import ImageExtractor
        self.image_extractor = ImageExtractor()
        
        # Create image converter
        def convert_image(image):
            return {
                "src": self.image_extractor.extract_image(image, context)
            }
        
        # Convert with mammoth
        with open(context.working_doc, 'rb') as docx:
            result = mammoth.convert_to_html(
                docx,
                id_prefix="footnote-",
                convert_image=mammoth.images.img_element(convert_image),
                ignore_empty_paragraphs=True
            )
            
            context.html_content = result.value
            
            if result.messages:
                for msg in result.messages:
                    print(f"  Mammoth: {msg}")
                    
        print(f"✓ HTML generated")
        print(f"✓ {len(context.images)} images extracted")
        
        return context