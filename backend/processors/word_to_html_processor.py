from pathlib import Path
import mammoth
from .base_processor import BaseProcessor

class WordToHtmlProcessor(BaseProcessor):
    """Convert Word documents to HTML using mammoth"""
    
    async def process(self, file_path: Path) -> dict:
        """Convert Word to HTML with proper formatting"""
        
        # Custom style mappings for better HTML
        style_map = """
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
        p[style-name='Title'] => h1.title
        p[style-name='Subtitle'] => h2.subtitle
        """
        
        # Convert with mammoth
        with open(file_path, "rb") as docx_file:
            result = mammoth.convert_to_html(
                docx_file,
                style_map=style_map,
                convert_image=mammoth.images.img_element(self._convert_image)
            )
        
        # Create full HTML document
        html_content = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{file_path.stem}</title>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }}
        h1, h2, h3 {{ color: #333; }}
        img {{ max-width: 100%; height: auto; }}
        table {{ border-collapse: collapse; width: 100%; }}
        td, th {{ border: 1px solid #ddd; padding: 8px; }}
    </style>
</head>
<body>
    {result.value}
</body>
</html>"""
        
        # Save HTML file
        output_path = file_path.parent / f"{file_path.stem}.html"
        output_path.write_text(html_content, encoding='utf-8')
        
        # Log any conversion messages
        if result.messages:
            print(f"Conversion messages for {file_path.name}:")
            for message in result.messages:
                print(f"  - {message}")
        
        return {
            "filename": file_path.name,
            "path": str(output_path),
            "messages": [str(m) for m in result.messages]
        }
    
    def _convert_image(self, image):
        """Handle image conversion"""
        import base64
        with image.open() as image_bytes:
            encoded = base64.b64encode(image_bytes.read()).decode('ascii')
        return {
            "src": f"data:{image.content_type};base64,{encoded}"
        }