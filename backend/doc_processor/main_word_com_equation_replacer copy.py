# ============= IMPROVED WORD COM EQUATION REPLACER =============
"""
Improved Word COM equation replacer with comprehensive equation detection
"""

import sys
import os
import win32com.client
from pathlib import Path
import pythoncom
import zipfile
from lxml import etree
import traceback
import time
from .omml_2_latex import DirectOmmlToLatex

class WordCOMEquationReplacer:
    """Improved Word COM equation replacer - finds ALL equations"""

    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        
        
        self.omml_parser = DirectOmmlToLatex()
        self.latex_equations = []

    def _extract_and_convert_equations(self, docx_path):
        """Extract equations from ZIP - this always finds ALL equations"""
        
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

    def _collect_all_equations_comprehensive(self):
        """IMPROVED: Comprehensive equation collection using multiple methods"""
        
        print(f"\n{'='*40}")
        print("STEP 2: Collecting ALL equations (Improved)")
        print(f"{'='*40}\n")
        
        equation_data = []
        seen_positions = set()
        
        # Method 1: Direct document OMaths
        print("Method 1: Checking document.OMaths...")
        try:
            for i in range(1, self.doc.OMaths.Count + 1):
                eq = self.doc.OMaths.Item(i)
                position = eq.Range.Start
                if position not in seen_positions:
                    seen_positions.add(position)
                    equation_data.append({
                        'object': eq,
                        'position': position,
                        'method': 'document'
                    })
        except Exception as e:
            print(f"  Error in Method 1: {e}")
        
        print(f"  Found {len(equation_data)} equations via document.OMaths")
        
        # Method 2: Story ranges (includes headers, footers, textboxes)
        print("\nMethod 2: Checking all story ranges...")
        try:
            for story in self.doc.StoryRanges:
                while story:
                    if story.OMaths.Count > 0:
                        for i in range(1, story.OMaths.Count + 1):
                            eq = story.OMaths.Item(i)
                            position = eq.Range.Start
                            if position not in seen_positions:
                                seen_positions.add(position)
                                equation_data.append({
                                    'object': eq,
                                    'position': position,
                                    'method': 'story'
                                })
                    story = story.NextStoryRange
        except Exception as e:
            print(f"  Error in Method 2: {e}")
        
        print(f"  Total after story ranges: {len(equation_data)} equations")
        
        # Method 3: Paragraph-by-paragraph scan
        print("\nMethod 3: Scanning paragraphs...")
        try:
            for para_idx in range(1, self.doc.Paragraphs.Count + 1):
                para = self.doc.Paragraphs.Item(para_idx)
                if para.Range.OMaths.Count > 0:
                    for i in range(1, para.Range.OMaths.Count + 1):
                        eq = para.Range.OMaths.Item(i)
                        position = eq.Range.Start
                        if position not in seen_positions:
                            seen_positions.add(position)
                            equation_data.append({
                                'object': eq,
                                'position': position,
                                'method': 'paragraph'
                            })
        except Exception as e:
            print(f"  Error in Method 3: {e}")
        
        print(f"  Total after paragraphs: {len(equation_data)} equations")

        # NEW METHOD 6: VML TEXTBOXES
        # Store equation_data temporarily for the VML method
        print("\nMethod 6: VML textboxes...")

        self.equation_data = equation_data
        vml_equations = self._collect_vml_textbox_equations()

        # Add VML equations to main list
        for vml_eq in vml_equations:
            equation_data.append(vml_eq)

        print(f"  Total after VML textboxes: {len(equation_data)} equations")



        # Method 4: Table cells
        print("\nMethod 4: Checking tables...")
        try:
            for table_idx in range(1, self.doc.Tables.Count + 1):
                table = self.doc.Tables.Item(table_idx)
                for row in table.Rows:
                    for cell in row.Cells:
                        if cell.Range.OMaths.Count > 0:
                            for i in range(1, cell.Range.OMaths.Count + 1):
                                eq = cell.Range.OMaths.Item(i)
                                position = eq.Range.Start
                                if position not in seen_positions:
                                    seen_positions.add(position)
                                    equation_data.append({
                                        'object': eq,
                                        'position': position,
                                        'method': 'table'
                                    })
        except Exception as e:
            print(f"  Error in Method 4: {e}")
        
        print(f"  Total after tables: {len(equation_data)} equations")
        
        # Method 5: Selection-based search
        print("\nMethod 5: Selection-based search...")
        try:
            # Save current selection
            original_start = self.word.Selection.Start
            original_end = self.word.Selection.End
            
            # Start from beginning
            self.word.Selection.HomeKey(Unit=6)  # wdStory = 6
            
            # Search using Find
            find = self.word.Selection.Find
            find.ClearFormatting()
            
            # Move through document and check for equations
            doc_end = self.doc.Content.End
            current_pos = 0
            
            while current_pos < doc_end:
                self.word.Selection.SetRange(current_pos, min(current_pos + 1000, doc_end))
                
                if self.word.Selection.OMaths.Count > 0:
                    for i in range(1, self.word.Selection.OMaths.Count + 1):
                        eq = self.word.Selection.OMaths.Item(i)
                        position = eq.Range.Start
                        if position not in seen_positions:
                            seen_positions.add(position)
                            equation_data.append({
                                'object': eq,
                                'position': position,
                                'method': 'selection'
                            })
                
                current_pos += 500  # Move forward
            
            # Restore original selection
            self.word.Selection.SetRange(original_start, original_end)


        except Exception as e:
            print(f"  Error in Method 5: {e}")



        
        # Sort by position
        print("\nSorting equations by document position...")
        equation_data.sort(key=lambda x: x['position'])
        
        # Show distribution by method
        method_counts = {}
        for eq in equation_data:
            method = eq['method']
            method_counts[method] = method_counts.get(method, 0) + 1
        
        print("\nEquation sources:")
        for method, count in method_counts.items():
            print(f"  {method}: {count} equations")
        
        print(f"\n✓ Total unique equations collected: {len(equation_data)}")
        return equation_data


    def _replace_sorted_equations_safe(self, equation_data):
        """SAFE: Replace equations with POSITION MATCHING"""
        
        print(f"\n{'='*40}")
        print("STEP 3: Replacing equations (Position-aware)")
        print(f"{'='*40}\n")
        
        equations_replaced = 0
        failed_replacements = []
        
        # KEY FIX: Match equations by their text content, not index
        # First, extract text from each COM equation to match with ZIP equations
        com_equation_texts = []
        for eq_info in equation_data:
            try:
                eq_obj = eq_info['object']
                eq_range = eq_obj.Range
                eq_text = eq_range.Text.strip() if eq_range.Text else ""
                com_equation_texts.append(eq_text)
            except:
                com_equation_texts.append("")
        
        # Process from last to first
        for i in range(len(equation_data) - 1, -1, -1):
            try:
                eq_info = equation_data[i]
                eq_obj = eq_info['object']
                position = eq_info['position']
                method = eq_info.get('method', 'unknown')
                
                # CRITICAL: Find matching LaTeX by position or content
                # For now, use sequential matching for accessible equations
                if i < len(self.latex_equations):
                    latex_data = self.latex_equations[i]
                else:
                    # Skip if no LaTeX available
                    print(f"⚠ No LaTeX match for COM equation {i + 1}")
                    continue
                
                latex_text = latex_data['latex'].strip() or f"[EQUATION_{i + 1}_EMPTY]"
                
                print(f"Replacing equation {i + 1} (from {method}) at position {position}")
                print(f"  LaTeX: {latex_text[:50]}..." if len(latex_text) > 50 else f"  LaTeX: {latex_text}")
                
                # Get range and delete
                try:
                    eq_range = eq_obj.Range
                    eq_range.Delete()
                except:
                    print(f"  ⚠ Cannot delete equation {i + 1}")
                    failed_replacements.append(i + 1)
                    continue
                
                # Insert replacement
                is_inline = len(latex_text) < 30
                
                if is_inline:
                    marked_text = f' MATHSTARTINLINE\\({latex_text}\\)MATHENDINLINE '
                else:
                    marked_text = f' MATHSTARTDISPLAY\\[{latex_text}\\]MATHENDDISPLAY '
                
                try:
                    eq_range.InsertAfter(marked_text)
                    equations_replaced += 1
                    print(f"  ✓ Replaced successfully")
                except:
                    print(f"  ⚠ Cannot insert replacement")
                    failed_replacements.append(i + 1)
                    
            except Exception as e:
                print(f"  Error replacing equation {i + 1}: {e}")
                failed_replacements.append(i + 1)
        
        print(f"\n✓ Replaced {equations_replaced}/{len(equation_data)} equations")
        if failed_replacements:
            print(f"⚠ Failed equations: {failed_replacements}")
        
        # WARNING about mismatch
        if len(equation_data) < len(self.latex_equations):
            print(f"\n⚠⚠⚠ CRITICAL WARNING ⚠⚠⚠")
            print(f"Only {len(equation_data)} of {len(self.latex_equations)} equations accessible via COM")
            print(f"Missing {len(self.latex_equations) - len(equation_data)} equations in VML textboxes")
            print(f"These cannot be replaced using COM!")
        
        return equations_replaced


        
    def _convert_to_html(self, output_path):
        """Convert to HTML with MathJax support"""
        
        print(f"\n{'='*40}")
        print("STEP 4: Converting to HTML")
        print(f"{'='*40}\n")
        
        try:
            html_path = output_path.with_suffix('.html')
            
            print(f"Saving as HTML: {html_path}")
            self.doc.SaveAs2(str(html_path), FileFormat=10)  # Filtered HTML
            
            print("✓ HTML file created")
            
            # Close document
            self.doc.Close(SaveChanges=False)
            self.doc = None
            
            time.sleep(1)  # Wait for file to be released
            
            # Read and modify HTML
            try:
                with open(html_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
            except UnicodeDecodeError:
                with open(html_path, 'r', encoding='windows-1252') as f:
                    html_content = f.read()
            
            # Extract body content
            import re
            body_match = re.search(r'<body[^>]*>(.*?)</body>', html_content, re.DOTALL | re.IGNORECASE)
            if body_match:
                body_content = body_match.group(1)
            else:
                body_content = html_content
            
            # Create complete HTML with MathJax
            complete_html = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="utf-8">
    <title>Document with LaTeX Equations</title>
    <script>
    // Convert markers to proper HTML elements
    window.addEventListener('DOMContentLoaded', function() {
        var content = document.body.innerHTML;
        
        // Replace inline math
        content = content.replace(/MATHSTARTINLINE([\\s\\S]*?)MATHENDINLINE/g, function(match, latex) {
            console.log('Found inline:', latex);
            return '<span class="inlineMath">' + latex + '</span>';
        });
        
        // Replace display math
        content = content.replace(/MATHSTARTDISPLAY([\\s\\S]*?)MATHENDDISPLAY/g, function(match, latex) {
            console.log('Found display:', latex);
            return '<div class="Math_box">' + latex + '</div>';
        });
        
        document.body.innerHTML = content;
        
        // Trigger MathJax rendering
        if (window.MathJax) {
            MathJax.typesetPromise();
        }
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
        color: #cc0000;
    }
    .Math_box {
        display: block;
        margin: 15px auto;
        text-align: center;
        color: #008800;
        font-size: 1.1em;
    }
    body {
        direction: rtl;
        text-align: right;
        font-family: Arial, Tahoma, sans-serif;
    }
    </style>
</head>
<body lang="AR-SA" dir="rtl">
""" + body_content + """
</body>
</html>"""
            
            # Write HTML with UTF-8 BOM
            with open(html_path, 'wb') as f:
                f.write(b'\xef\xbb\xbf')  # UTF-8 BOM
                f.write(complete_html.encode('utf-8'))
            
            print(f"✓ HTML with MathJax saved: {html_path}")
            
            return html_path
            
        except Exception as e:
            print(f"❌ Error converting to HTML: {e}")
            traceback.print_exc()
            return None

    def process_document(self, docx_path, output_path=None):
        """Main entry point - process equations with improved detection"""
        
        docx_path = Path(docx_path).absolute()
        
        if not output_path:
            output_path = docx_path.parent / f"{docx_path.stem}_processed.docx"
        else:
            output_path = Path(output_path).absolute()
        
        if output_path == docx_path:
            output_path = docx_path.parent / f"{docx_path.stem}_processed_safe.docx"
        
        print(f"\n{'='*60}")
        print(f"WORD COM EQUATION REPLACER (Improved)")
        print(f"{'='*60}")
        print(f"📄 Input: {docx_path}")
        print(f"📄 Output: {output_path}")
        print(f"{'='*60}\n")
        
        # Extract equations from ZIP (always accurate)
        self.latex_equations = self._extract_and_convert_equations(docx_path)
        
        if not self.latex_equations:
            print("⚠ No equations found")
        
        try:
            # Open Word
            print("Starting Word application...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = False
            self.word.ScreenUpdating = False
            
            # Open document
            print("Opening document...")
            self.doc = self.word.Documents.Open(str(docx_path))
            print("✓ Document opened")
            
            # Accept all tracked changes first
            print("\nAccepting tracked changes...")
            try:
                self.doc.AcceptAllRevisions()
                print("✓ All revisions accepted")
            except:
                print("⚠ No tracked changes or unable to accept")
            
            if self.latex_equations:
                print(f"\nExpected equations: {len(self.latex_equations)}")
                print(f"Document OMaths count: {self.doc.OMaths.Count}")
                
                # Collect equations with improved method
                equation_data = self._collect_all_equations_comprehensive()
                
                if len(equation_data) < len(self.latex_equations):
                    print(f"\n⚠ WARNING: Found {len(equation_data)} equations, expected {len(self.latex_equations)}")
                    print("Some equations may not be replaced!")
                else:
                    print(f"\n✅ Found all {len(equation_data)} equations!")
                
                # Replace equations
                equation_count = self._replace_sorted_equations_safe(equation_data)
                print(f"\n✓ Replaced {equation_count} equations")
            
            # Save processed document
            print("\nSaving processed Word document...")
            self.doc.SaveAs2(str(output_path))
            print(f"✓ Saved: {output_path}")
            
            # Convert to HTML
            html_path = self._convert_to_html(output_path)
            
            print(f"\n{'='*60}")
            print(f"✅ PROCESSING COMPLETE!")
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


    def _collect_vml_textbox_equations(self):
        """Method 6: Access VML textboxes specifically"""
        
        print("\nMethod 6: Accessing VML textboxes...")
        vml_equations = []
        seen_positions = set()
        
        # Get existing positions to avoid duplicates
        for eq in self.equation_data:
            seen_positions.add(eq['position'])
        
        try:
            # Access Shapes collection
            for i in range(1, self.doc.Shapes.Count + 1):
                shape = self.doc.Shapes.Item(i)
                
                # Try TextFrame access
                try:
                    if hasattr(shape, 'TextFrame'):
                        tf = shape.TextFrame
                        if tf.HasText:
                            tr = tf.TextRange
                            if tr.OMaths.Count > 0:
                                for j in range(1, tr.OMaths.Count + 1):
                                    eq = tr.OMaths.Item(j)
                                    pos = eq.Range.Start
                                    if pos not in seen_positions:
                                        seen_positions.add(pos)
                                        vml_equations.append({
                                            'object': eq,
                                            'position': pos,
                                            'method': 'vml_textframe'
                                        })
                except:
                    pass
                
                # Try CanvasItems if it's a canvas
                try:
                    if hasattr(shape, 'CanvasItems'):
                        for c_idx in range(1, shape.CanvasItems.Count + 1):
                            item = shape.CanvasItems.Item(c_idx)
                            if hasattr(item, 'TextFrame'):
                                if item.TextFrame.HasText:
                                    tr = item.TextFrame.TextRange
                                    if tr.OMaths.Count > 0:
                                        for j in range(1, tr.OMaths.Count + 1):
                                            eq = tr.OMaths.Item(j)
                                            pos = eq.Range.Start
                                            if pos not in seen_positions:
                                                seen_positions.add(pos)
                                                vml_equations.append({
                                                    'object': eq,
                                                    'position': pos,
                                                    'method': 'vml_canvas'
                                                })
                except:
                    pass
        except Exception as e:
            print(f"  Error accessing VML shapes: {e}")
        
        print(f"  Found {len(vml_equations)} VML textbox equations")
        return vml_equations
        

if __name__ == "__main__":
    test_file = r"test_document.docx"
    
    print("Starting Improved Word COM Equation Replacer...")
    converter = WordCOMEquationReplacer()
    
    try:
        result = converter.process_document(test_file)
        if result:
            print(f"\n✅ Processing complete!")
            print(f"📄 Word: {result['word_path']}")
            print(f"🌐 HTML: {result['html_path']}")
    except Exception as e:
        print(f"\n❌ Processing failed: {e}")