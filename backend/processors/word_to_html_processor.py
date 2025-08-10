"""
FILE 1: backend/processors/word_to_html_processor.py
This processor actually creates HTML files
"""

from pathlib import Path
import mammoth
from .base_processor import BaseProcessor
import logging

logger = logging.getLogger(__name__)

class WordToHtmlProcessor(BaseProcessor):
    """Convert Word documents to HTML using mammoth"""
    
    async def process(self, file_path: Path, output_dir: Path) -> dict:
        """Convert Word to HTML with proper formatting"""
        
        logger.info(f"Converting {file_path.name} to HTML")
        
        # Custom style mappings for better HTML
        style_map = """
        p[style-name='Heading 1'] => h1:fresh
        p[style-name='Heading 2'] => h2:fresh
        p[style-name='Heading 3'] => h3:fresh
        p[style-name='Title'] => h1.title
        p[style-name='Subtitle'] => h2.subtitle
        """
        
        try:
            # Convert with mammoth
            with open(file_path, "rb") as docx_file:
                result = mammoth.convert_to_html(
                    docx_file,
                    style_map=style_map,
                    convert_image=mammoth.images.img_element(self._convert_image)
                )
            
            # Create full HTML document
            html_content = f"""<!DOCTYPE html>
<html lang="ar" xx="3" dir="rtl">
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
            direction: rtl;
        }}
        h1, h2, h3 {{ 
            color: #333; 
            margin-top: 1.5em;
            margin-bottom: 0.5em;
        }}
        h1 {{ font-size: 2em; }}
        h2 {{ font-size: 1.5em; }}
        h3 {{ font-size: 1.2em; }}
        p {{ 
            margin: 1em 0;
            text-align: justify;
        }}
        img {{ 
            max-width: 100%; 
            height: auto; 
            display: block;
            margin: 1em auto;
        }}
        table {{ 
            border-collapse: collapse; 
            width: 100%;
            margin: 1em 0;
        }}
        td, th {{ 
            border: 1px solid #ddd; 
            padding: 8px;
            text-align: right;
        }}
        th {{
            background-color: #f2f2f2;
            font-weight: bold;
        }}
        ul, ol {{
            margin: 1em 0;
            padding-right: 2em;
        }}
        li {{
            margin: 0.5em 0;
        }}
    </style>
</head>
<body>
    {result.value}
</body>
</html>"""
            
            # Save HTML file
            output_filename = f"{file_path.stem}.html"
            output_path = output_dir / output_filename
            output_path.write_text(html_content, encoding='utf-8')
            
            logger.info(f"HTML saved to: {output_path}")
            
            # Log any conversion messages
            if result.messages:
                logger.warning(f"Conversion messages for {file_path.name}:")
                for message in result.messages:
                    logger.warning(f"  - {message}")
            
            return {
                "filename": file_path.name,
                "output_filename": output_filename,
                "path": str(output_path),
                "messages": [str(m) for m in result.messages] if result.messages else [],
                "success": True
            }
            
        except Exception as e:
            logger.error(f"Failed to convert {file_path.name}: {str(e)}")
            raise
    
    def _convert_image(self, image):
        """Handle image conversion to base64"""
        import base64
        try:
            with image.open() as image_bytes:
                encoded = base64.b64encode(image_bytes.read()).decode('ascii')
            return {
                "src": f"data:{image.content_type};base64,{encoded}"
            }
        except Exception as e:
            logger.error(f"Failed to convert image: {str(e)}")
            return {"src": ""}

