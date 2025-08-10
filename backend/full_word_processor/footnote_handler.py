# backend/processors/full-word-processor/footnote_handler.py

from .models import ProcessingContext

class FootnoteHandler:
    """Handles footnote anchor generation and linking"""
    
    def process(self, context: ProcessingContext) -> ProcessingContext:
        """Process footnotes and generate anchors"""
        
        print("\nSTEP 3: Processing footnotes...")
        
        if not context.footnotes:
            print("✓ No footnotes to process")
            return context
            
        # Generate anchors for each footnote
        for footnote in context.footnotes:
            footnote.anchor = self._generate_anchor(footnote.id)
            
        print(f"✓ Generated anchors for {len(context.footnotes)} footnotes")
        
        return context
    
    def _generate_anchor(self, footnote_id):
        """Generate unique anchor for footnote"""
        return f"fn-{footnote_id}"