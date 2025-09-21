# ============= COMPLETE WORD TO HTML CONVERTER =============
"""Process Word document equations and convert to HTML"""

import sys
import os
import win32com.client
from pathlib import Path
import pythoncom
import json
import zipfile
from lxml import etree
import traceback
import shutil

sys.path.append(str(Path(__file__).parent.parent))

# RESTORED FROM OLD CODE
from core.logger import setup_logger
logger = setup_logger("word_to_html_converter")

try:
    from doc_processor.omml_2_latex import DirectOmmlToLatex
except ImportError:
    from omml_2_latex import DirectOmmlToLatex


#class WordToHTMLConverter:
class WordCOMEquationReplacer:

    """Complete Word to HTML conversion with equation processing"""

    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        self.omml_parser = DirectOmmlToLatex()
        self.latex_equations = []

    def _extract_and_convert_equations(self, docx_path):
        """Extract equations from ZIP"""
        
        print(f"\n{'='*40}")
        print("STEP 1: Extracting equations from ZIP")
        print(f"{'='*40}")
        
        results = []
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as z:
                with z.open('word/document.xml') as f:
                    content = f.read()
                    root = etree.fromstring(content)
                    
                    ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
                    equations = root.xpath('//m:oMath', namespaces=ns)
                    
                    print(f"Found {len(equations)} equations in XML\n")
                    
                    for i, eq in enumerate(equations, 1):
                        texts = eq.xpath('.//m:t/text()', namespaces=ns)
                        text = ''.join(texts)
                        latex = self.omml_parser.parse(eq)
                        
                        results.append({
                            'index': i,
                            'text': text,
                            'latex': latex
                        })
                        
                        print(f"  Equation {i}: {latex[:50]}..." if len(latex) > 50 else f"  Equation {i}: {latex}")
            
            print(f"\n✓ Successfully converted {len(results)} equations")
            return results
            
        except Exception as e:
            print(f"❌ Error extracting equations: {e}")
            traceback.print_exc()
            return []

    def _collect_all_equations_old(self):
        """Collect ALL equations with positions and sort them"""
        
        print(f"\n{'='*40}")
        print("STEP 2: Collecting ALL equations with positions")
        print(f"{'='*40}\n")
        
        equation_data = []
        
        print("Using Find method to locate equations...")
        try:
            # Start from beginning
            self.word.Selection.HomeKey(Unit=6)  # wdStory = 6
            
            # Track unique positions
            seen_positions = set()
            
            # Search through entire document
            while True:
                # Move through document
                self.word.Selection.MoveRight()
                
                # Check for equations
                if self.word.Selection.OMaths.Count > 0:
                    for i in range(1, self.word.Selection.OMaths.Count + 1):
                        try:
                            eq = self.word.Selection.OMaths.Item(i)
                            
                            # Get position
                            position = eq.Range.Start
                            
                            # Only add if not seen before
                            if position not in seen_positions:
                                seen_positions.add(position)
                                equation_data.append({
                                    'object': eq,
                                    'position': position
                                })
                                
                                if len(equation_data) % 10 == 0:
                                    print(f"  Found {len(equation_data)} equations...")
                        except:
                            pass
                
                # Stop conditions
                if self.word.Selection.End >= self.doc.Content.End - 1:
                    break
                
                if len(equation_data) >= 144:
                    break
            
            print(f"  Found {len(equation_data)} equations via Find")
            
        except Exception as e:
            print(f"  Find method error: {e}")
        
        # Sort by position
        print("\nSorting equations by document position...")
        equation_data.sort(key=lambda x: x['position'])
        
        print(f"\n✓ Collected and sorted {len(equation_data)} equations")
        return equation_data



    def _collect_all_equations(self):
        """Collect ALL equations with positions and sort them"""
        
        print(f"\n{'='*40}")
        print("STEP 2: Collecting ALL equations with positions")
        print(f"{'='*40}\n")
        
        equation_data = []
        seen_positions = set()
        
        print("Searching entire document comprehensively...")
        
        # Method 1: Direct OMaths collection
        try:
            for i in range(1, self.doc.OMaths.Count + 1):
                eq = self.doc.OMaths.Item(i)
                position = eq.Range.Start
                if position not in seen_positions:
                    seen_positions.add(position)
                    equation_data.append({
                        'object': eq,
                        'position': position
                    })
        except:
            pass
        
        # Method 2: Search through entire document range
        try:
            doc_range = self.doc.Content
            if doc_range.OMaths.Count > 0:
                for i in range(1, doc_range.OMaths.Count + 1):
                    eq = doc_range.OMaths.Item(i)
                    position = eq.Range.Start
                    if position not in seen_positions:
                        seen_positions.add(position)
                        equation_data.append({
                            'object': eq,
                            'position': position
                        })
        except:
            pass
        
        # Method 3: Iterate through all ranges
        for rng in self.doc.Range().Paragraphs:
            try:
                if rng.Range.OMaths.Count > 0:
                    for i in range(1, rng.Range.OMaths.Count + 1):
                        eq = rng.Range.OMaths.Item(i)
                        position = eq.Range.Start
                        if position not in seen_positions:
                            seen_positions.add(position)
                            equation_data.append({
                                'object': eq,
                                'position': position
                            })
            except:
                pass
        
        print(f"Found {len(equation_data)} unique equations")
        
        # Sort by position
        equation_data.sort(key=lambda x: x['position'])
        
        print(f"\n✓ Collected and sorted {len(equation_data)} equations")
        return equation_data


    def _replace_sorted_equations_old(self, equation_data):
        """Replace equations with LaTeX markers"""
        
        print(f"\n{'='*40}")
        print("STEP 3: Replacing equations with LaTeX markers")
        print(f"{'='*40}\n")
        
        equations_replaced = 0
        
        for i in range(len(equation_data) - 1, -1, -1):
            if i >= len(self.latex_equations):
                continue
            
            try:
                eq_info = equation_data[i]
                eq_obj = eq_info['object']
                position = eq_info['position']
                
                # Get LaTeX
                latex_data = self.latex_equations[i]
                latex_text = latex_data['latex'].strip() or f"[EQUATION_{i + 1}_EMPTY]"
                
                print(f"Replacing equation {i + 1} at position {position}: {latex_text[:30]}...")
                
                # Replace
                eq_range = eq_obj.Range
                eq_range.Delete()
                
                # Use text markers that Word won't remove
                is_inline = len(latex_text) < 30
                
                if is_inline:
                    marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
                else:
                    marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '
                
                eq_range.InsertAfter(marked_text)
                equations_replaced += 1
                print(f"  ✓ Replaced")
                
            except Exception as e:
                print(f"  Error replacing equation {i + 1}: {e}")
        
        return equations_replaced

    def _replace_sorted_equations_old2(self, equation_data):
        """Replace equations with LaTeX markers"""
        
        print(f"\n{'='*40}")
        print("STEP 3: Replacing equations with LaTeX markers")
        print(f"{'='*40}\n")
        
        equations_replaced = 0
        
        # Process from last to first
        for i in range(len(equation_data) - 1, -1, -1):
            if i >= len(self.latex_equations):
                continue
            
            try:
                eq_info = equation_data[i]
                eq_obj = eq_info['object']
                position = eq_info['position']
                
                # Get LaTeX
                latex_data = self.latex_equations[i]
                latex_text = latex_data['latex'].strip() or f"[EQUATION_{i + 1}_EMPTY]"
                
                print(f"Replacing equation {i + 1} at position {position}: {latex_text[:30]}...")
                
                # Replace
                eq_range = eq_obj.Range
                eq_range.Delete()
                
                # Insert with markers for HTML
                is_inline = len(latex_text) < 30
                if is_inline:
                    marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
                else:
                    marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '
                
                eq_range.InsertAfter(f" {marked_text} ")
                equations_replaced += 1
                
            except Exception as e:
                print(f"  Error replacing equation {i + 1}: {e}")
        
        return equations_replaced

    def _replace_sorted_equations(self, equation_data):
        """Replace equations with LaTeX markers - matching by content"""
        
        print(f"\n{'='*40}")
        print("STEP 3: Replacing equations with LaTeX markers")
        print(f"{'='*40}\n")
        
        equations_replaced = 0
        
        # Process from last to first
        for i in range(len(equation_data) - 1, -1, -1):
            try:
                eq_info = equation_data[i]
                eq_obj = eq_info['object']
                position = eq_info['position']
                
                # Get the equation's text content from Word
                eq_text = eq_obj.Range.Text.strip()
                
                # Find matching LaTeX by content, not by index
                latex_text = None
                for latex_data in self.latex_equations:
                    # Try to match by text content
                    if latex_data['text'].strip() == eq_text or \
                    latex_data['latex'].strip() == eq_text:
                        latex_text = latex_data['latex'].strip()
                        break
                
                # If no match found, use a fallback
                if not latex_text:
                    # Use index if within bounds
                    if i < len(self.latex_equations):
                        latex_text = self.latex_equations[i]['latex'].strip()
                    else:
                        latex_text = f"[EQUATION_{i + 1}_NO_MATCH]"
                
                print(f"Replacing equation {i + 1} at position {position}: {latex_text[:30]}...")
                
                # Replace
                eq_range = eq_obj.Range
                eq_range.Delete()
                
                # Insert with markers
                is_inline = len(latex_text) < 30
                if is_inline:
                    marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
                else:
                    marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '
                
                eq_range.InsertAfter(marked_text)
                equations_replaced += 1
                
            except Exception as e:
                print(f"  Error replacing equation {i + 1}: {e}")
        
        return equations_replaced


    def _convert_to_html(self, output_path):
        """Convert the processed Word document to HTML"""
        
        print(f"\n{'='*40}")
        print("STEP 4: Converting to HTML")
        print(f"{'='*40}\n")
        
        try:
            # Save as HTML using Word COM
            html_path = output_path.with_suffix('.html')
            
            print(f"Saving as HTML: {html_path}")
            self.doc.SaveAs2(str(html_path), FileFormat=10)  # Filtered HTML
            
            print("✓ HTML file created")
            
            # IMPORTANT: Close the document to release the HTML file
            self.doc.Close(SaveChanges=False)
            self.doc = None
            
            # Small delay to ensure file is released
            import time
            time.sleep(1)
            
            # Read the HTML content
            try:
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
            except UnicodeDecodeError:
                with open(html_path, 'r', encoding='windows-1252') as f:
                    html_content = f.read()
            
            # Extract just the body content from Word's HTML
            import re
            body_match = re.search(r'<body[^>]*>(.*?)</body>', html_content, re.DOTALL | re.IGNORECASE)
            if body_match:
                body_content = body_match.group(1)
            else:
                body_content = html_content
            
            # Create complete HTML with your exact structure
            complete_html = """<!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
    <meta http-equiv=Content-Type content="text/html; charset=utf-8">
    <meta name=Generator content="Microsoft Word 15 (filtered)">
    <meta charset="utf-8">
    <script>
    window.addEventListener('DOMContentLoaded', function() {
        var content = document.body.innerHTML;
        
        // Debug - check what we have
        console.log('Sample before:', content.substring(0, 500));
        
        // Replace inline math - handles any content between markers
        content = content.replace(/MATHSTARTINLINE([\\s\\S]*?)MATHENDINLINE/g, function(match, latex) {
            console.log('Found inline:', latex);
            return '<span class="inlineMath">' + latex + '</span>';
        });
        
        // Replace display math - handles any content between markers
        content = content.replace(/MATHSTARTDISPLAY([\\s\\S]*?)MATHENDDISPLAY/g, function(match, latex) {
            console.log('Found display:', latex);
            return '<div class="Math_box">' + latex + '</div>';
        });
        
        document.body.innerHTML = content;
        
        console.log('Replacement done!');
    });
    </script>
    <script>
    window.MathJax = {
        tex: {
            inlineMath: [['\\\\(', '\\\\)']],
            displayMath: [['\\\\[', '\\\\]']]
        }
    };
    </script>
    <script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-chtml.js"></script>
    <style>
    .inlineMath {
        display: inline;
        margin: 0 2px;
        color: red;
    }
    .Math_box {
        display: block;
        margin: 15px auto;
        text-align: center;
        color: green;
    }
    body {
        direction: rtl;
        text-align: right;
        font-family: Arial, Tahoma, sans-serif;
    }
    </style>
    </head>
    <body lang=AR-SA dir="rtl">
    """ + body_content + """
    </body>
    </html>"""
            
            # Write the complete HTML with UTF-8 BOM
            with open(html_path, 'wb') as f:
                f.write(b'\xef\xbb\xbf')  # UTF-8 BOM for Arabic
                f.write(complete_html.encode('utf-8'))
            
            print(f"✓ HTML with MathJax saved: {html_path}")
            
            return html_path
            
        except Exception as e:
            print(f"❌ Error converting to HTML: {e}")
            import traceback
            traceback.print_exc()
            return None


    def process_document(self, docx_path, output_path=None):
        """Main entry point - process equations and convert to HTML"""
        
        docx_path = Path(docx_path).absolute()
        
        if not output_path:
            output_path = docx_path.parent / f"{docx_path.stem}_processed.docx"
        else:
            output_path = Path(output_path).absolute()
        
        if output_path == docx_path:
            output_path = docx_path.parent / f"{docx_path.stem}_processed_safe.docx"
        
        print(f"\n{'='*60}")
        print(f"COMPLETE WORD TO HTML CONVERSION")
        print(f"{'='*60}")
        print(f"📍 Input: {docx_path}")
        print(f"📍 Output: {output_path}")
        print(f"{'='*60}\n")
        
        # Extract equations from ZIP
        self.latex_equations = self._extract_and_convert_equations(docx_path)
        
        if not self.latex_equations:
            print("⚠ No equations found, will convert directly to HTML")
        
        try:
            # Open Word COM
            print("Starting Word application...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = False
            self.word.ScreenUpdating = False
            
            print("Opening document...")
            self.doc = self.word.Documents.Open(str(docx_path))
            print("✓ Document opened")
            
            if self.latex_equations:
                print(f"\nInitial doc.OMaths.Count: {self.doc.OMaths.Count}")
                
                # Find and replace equations
                equation_data = self._collect_all_equations()
                
                if len(equation_data) < len(self.latex_equations):
                    print(f"\n⚠ WARNING: Found {len(equation_data)} equations, expected {len(self.latex_equations)}")
                else:
                    print(f"\n✅ Found all {len(equation_data)} equations!")
                
                # Replace equations
                equation_count = self._replace_sorted_equations(equation_data)
                print(f"\n✓ Replaced {equation_count} equations")
            
            # Save processed Word document
            print("\nSaving processed Word document...")
            self.doc.SaveAs2(str(output_path))
            print(f"✓ Saved: {output_path}")
            
            # Convert to HTML
            html_path = self._convert_to_html(output_path)
            
            print(f"\n{'='*60}")
            print(f"✅ CONVERSION COMPLETE!")
            print(f"📄 Word output: {output_path}")
            if html_path:
                print(f"🌐 HTML output: {html_path}")
            print(f"{'='*60}\n")
            
            return {
                'word_path': output_path,
                'html_path': html_path
            }
            
        except Exception as e:
            print(f"\n❌ ERROR: {e}")
            traceback.print_exc()
            raise
            
        finally:
            self._cleanup()

    def _cleanup(self):
        """Clean up Word application"""
        try:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
        except:
            pass
        finally:
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    test_file = r"test_document.docx"
    
    print("Starting Complete Word to HTML Converter...")
    converter = WordToHTMLConverter()
    
    try:
        result = converter.process_document(test_file)
        if result:
            print(f"\n✅ Conversion complete!")
            print(f"📄 Word: {result['word_path']}")
            print(f"🌐 HTML: {result['html_path']}")
    except Exception as e:
        print(f"\n❌ Conversion failed: {e}")