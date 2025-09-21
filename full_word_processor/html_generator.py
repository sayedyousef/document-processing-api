# backend/processors/full-word-processor/html_generator.py

from pathlib import Path
from .models import ProcessingContext

class HTMLGenerator:
    """Generates final HTML with all components"""
    
    def generate(self, context: ProcessingContext) -> Path:
        """Generate final HTML file"""
        
        print("\nSTEP 6: Generating final HTML...")
        
        # Build HTML
        html = self._build_html(
            context.html_content,
            context.input_path.stem,
            context.footnotes
        )
        
        # Save file
        output_file = context.output_dir / f"{context.input_path.stem}.html"
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(html)
            
        print(f"✓ HTML saved: {output_file.name}")
        
        return output_file
    
    def _build_html(self, content, title, footnotes):
        """Build complete HTML document"""
        
        # Process footnotes in content
        if footnotes:
            content = self._process_footnote_references(content, footnotes)
            content += self._build_footnotes_section(footnotes)
        
        return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    
    <!-- MathJax for equations -->
    <script>
      window.MathJax = {{
        tex: {{
          inlineMath: [['\\\\(', '\\\\)']],
          displayMath: [['\\\\[', '\\\\]']]
        }}
      }};
    </script>
    <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js"></script>
</head>
<body>
{content}
</body>
</html>"""
    
    def _process_footnote_references(self, content, footnotes):
        """Add anchors to footnote references"""
        # Implementation for processing footnote references
        return content
    
    def _build_footnotes_section(self, footnotes):
        """Build footnotes section at end of document"""
        
        if not footnotes:
            return ""
            
        html = '\n<hr>\n<section class="footnotes">\n'
        
        for fn in footnotes:
            html += f'<p id="{fn.anchor}">'
            html += f'<sup>{fn.id}</sup> {fn.content}'
            html += f' <a href="#{fn.reference_id}">↩</a>'
            html += '</p>\n'
            
        html += '</section>\n'
        
        return html