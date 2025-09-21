# backend/processors/full-word-processor/document_extractor.py

import zipfile
from lxml import etree
from .models import ProcessingContext, Footnote

class DocumentExtractor:
    """Extracts document structure including footnotes"""
    
    def extract(self, context: ProcessingContext) -> ProcessingContext:
        """Extract document structure"""
        
        print("STEP 1: Extracting document structure...")
        
        try:
            with zipfile.ZipFile(context.working_doc, 'r') as z:
                # List all files in the docx
                files = z.namelist()
                
                # Extract main document info
                if 'word/document.xml' in files:
                    context.metadata['has_main_document'] = True
                
                # Extract footnotes
                if 'word/footnotes.xml' in files:
                    context.footnotes = self._extract_footnotes(z)
                    print(f"  ✓ Found {len(context.footnotes)} footnotes")
                else:
                    print(f"  ✓ No footnotes in document")
                
                # Check for endnotes
                if 'word/endnotes.xml' in files:
                    context.metadata['has_endnotes'] = True
                    print(f"  ✓ Document has endnotes")
                
                # Store metadata
                context.metadata['has_footnotes'] = len(context.footnotes) > 0
                
        except Exception as e:
            print(f"  ⚠ Document extraction failed: {e}")
            
        return context
    
    def _extract_footnotes(self, zipfile_obj):
        """Extract footnotes from document"""
        
        footnotes = []
        
        with zipfile_obj.open('word/footnotes.xml') as f:
            tree = etree.parse(f)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            # Find all footnotes
            for footnote in tree.xpath('//w:footnote', namespaces=ns):
                footnote_id = footnote.get('{%s}id' % ns['w'])
                
                # Skip separator/continuation footnotes
                if footnote_id in ['0', '-1']:
                    continue
                    
                # Extract text content
                texts = footnote.xpath('.//w:t/text()', namespaces=ns)
                content = ''.join(texts)
                
                if content:  # Only add non-empty footnotes
                    footnotes.append(Footnote(
                        id=footnote_id,
                        reference_id=f"footnote-ref-{footnote_id}",
                        content=content
                    ))
                
        return footnotes