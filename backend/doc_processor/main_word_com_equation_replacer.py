# ============= FIXED WORD COM EQUATION REPLACER =============
"""
Word COM equation replacer with proper equation mapping
Handles the reality that VML textbox equations are inaccessible
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
    """Word COM equation replacer with VML awareness"""

    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        self.omml_parser = DirectOmmlToLatex()
        self.latex_equations = []
        self.xml_equation_locations = []  # Track where equations are in XML

    def _extract_and_analyze_equations(self, docx_path):
        """Extract equations and analyze their locations in XML"""
        
        print(f"\n{'='*40}")
        print("STEP 1: Extracting and analyzing equations from ZIP")
        print(f"{'='*40}")
        
        results = []
        locations = []
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as z:
                with z.open('word/document.xml') as f:
                    content = f.read()
                    root = etree.fromstring(content)
                    
                    ns = {
                        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                        'v': 'urn:schemas-microsoft-com:vml'
                    }
                    
                    equations = root.xpath('//m:oMath', namespaces=ns)
                    
                    print(f"Found {len(equations)} equations in XML\n")
                    
                    for i, eq in enumerate(equations, 1):
                        # Get equation content
                        texts = eq.xpath('.//m:t/text()', namespaces=ns)
                        text = ''.join(texts)
                        latex = self.omml_parser.parse(eq)
                        
                        # Analyze location
                        parent_chain = []
                        current = eq
                        in_vml = False
                        
                        for _ in range(10):
                            parent = current.getparent()
                            if parent is None:
                                break
                            tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                            parent_chain.append(tag)
                            
                            # Check if in VML/Fallback
                            if tag in ['txbxContent', 'textbox', 'Fallback']:
                                in_vml = True
                            
                            current = parent
                        
                        location_type = 'vml_textbox' if in_vml else 'accessible'
                        
                        results.append({
                            'index': i,
                            'text': text,
                            'latex': latex
                        })
                        
                        locations.append({
                            'index': i,
                            'type': location_type,
                            'text': text[:20] if text else '[empty]'
                        })
                        
                        if location_type == 'vml_textbox':
                            print(f"  Equation {i}: IN VML TEXTBOX (inaccessible)")
                        else:
                            print(f"  Equation {i}: {latex[:50]}..." if len(latex) > 50 else f"  Equation {i}: {latex}")
            
            # Count accessible vs inaccessible
            accessible_count = sum(1 for loc in locations if loc['type'] == 'accessible')
            vml_count = sum(1 for loc in locations if loc['type'] == 'vml_textbox')
            
            print(f"\n📊 Equation Analysis:")
            print(f"  ✓ Accessible equations: {accessible_count}")
            print(f"  ❌ VML textbox equations (inaccessible): {vml_count}")
            print(f"  Total: {len(results)}")
            
            return results, locations
            
        except Exception as e:
            print(f"❌ Error extracting equations: {e}")
            traceback.print_exc()
            return [], []

    def _collect_accessible_equations(self):
        """Collect only the equations that COM can actually access"""
        
        print(f"\n{'='*40}")
        print("STEP 2: Collecting accessible equations via COM")
        print(f"{'='*40}\n")
        
        equation_data = []
        seen_positions = set()
        
        # Method 1: Direct document OMaths
        print("Collecting from main document...")
        try:
            for i in range(1, self.doc.OMaths.Count + 1):
                eq = self.doc.OMaths.Item(i)
                position = eq.Range.Start
                if position not in seen_positions:
                    seen_positions.add(position)
                    
                    # Try to get equation text for matching
                    eq_text = ""
                    try:
                        eq_text = eq.Range.Text or ""
                    except:
                        pass
                    
                    equation_data.append({
                        'object': eq,
                        'position': position,
                        'method': 'document',
                        'text': eq_text
                    })
        except Exception as e:
            print(f"  Error: {e}")
        
        print(f"  Found {len(equation_data)} equations via document.OMaths")
        
        # Method 2: Story ranges
        print("\nChecking story ranges...")
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
                                    'method': 'story',
                                    'text': ''
                                })
                    story = story.NextStoryRange
        except:
            pass
        
        print(f"  Total equations found: {len(equation_data)}")
        
        # Sort by position
        equation_data.sort(key=lambda x: x['position'])
        
        print(f"\n✓ Total accessible equations: {len(equation_data)}")
        return equation_data

    def _map_equations_smart(self, com_equations, xml_locations):
        """Create smart mapping between COM equations and XML equations"""
        
        print(f"\n{'='*40}")
        print("Creating equation mapping...")
        print(f"{'='*40}\n")
        
        # Get indices of accessible equations from XML analysis
        accessible_indices = []
        for loc in xml_locations:
            if loc['type'] == 'accessible':
                accessible_indices.append(loc['index'] - 1)  # Convert to 0-based
        
        print(f"Accessible equation indices from XML: {accessible_indices[:10]}...")
        print(f"COM found {len(com_equations)} equations")
        print(f"XML has {len(accessible_indices)} accessible equations")
        
        # Create mapping
        mapping = {}
        for i, com_eq in enumerate(com_equations):
            if i < len(accessible_indices):
                xml_idx = accessible_indices[i]
                mapping[i] = xml_idx
                print(f"  COM equation {i+1} -> XML equation {xml_idx+1}")
            else:
                print(f"  COM equation {i+1} -> No XML match")
        
        return mapping

    def _replace_equations_with_mapping(self, equation_data, mapping):
        """Replace equations using proper mapping"""
        
        print(f"\n{'='*40}")
        print("STEP 3: Replacing equations with correct mapping")
        print(f"{'='*40}\n")
        
        equations_replaced = 0
        failed_replacements = []
        
        # Process from last to first to maintain positions
        for i in range(len(equation_data) - 1, -1, -1):
            try:
                eq_info = equation_data[i]
                eq_obj = eq_info['object']
                position = eq_info['position']
                method = eq_info.get('method', 'unknown')
                
                # Get correct LaTeX using mapping
                xml_idx = mapping.get(i, -1)
                if xml_idx == -1 or xml_idx >= len(self.latex_equations):
                    print(f"⚠ No mapping for COM equation {i + 1}")
                    continue
                
                latex_data = self.latex_equations[xml_idx]
                latex_text = latex_data['latex'].strip() or f"[EQUATION_{xml_idx + 1}_EMPTY]"
                
                print(f"Replacing COM equation {i+1} with XML equation {xml_idx+1}")
                print(f"  Position: {position}, Method: {method}")
                print(f"  LaTeX: {latex_text[:50]}..." if len(latex_text) > 50 else f"  LaTeX: {latex_text}")
                
                # Get range and delete
                try:
                    eq_range = eq_obj.Range
                    eq_range.Delete()
                except:
                    print(f"  ⚠ Cannot delete equation")
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
        
        print(f"\n✓ Successfully replaced {equations_replaced}/{len(equation_data)} equations")
        
        # Final warning
        vml_count = len(self.latex_equations) - len(equation_data)
        if vml_count > 0:
            print(f"\n⚠️ IMPORTANT: {vml_count} equations in VML textboxes could not be replaced")
            print("  These equations are inaccessible via Word COM API")
            print("  They remain unchanged in the document")
        
        return equations_replaced

    def _convert_to_html(self, output_path):
        """Convert to HTML with MathJax support"""
        
        print(f"\n{'='*40}")
        print("STEP 4: Converting to HTML")
        print(f"{'='*40}\n")
        
        try:
            html_path = output_path.with_suffix('.html')
            
            print(f"Saving as HTML: {html_path}")
            self.doc.SaveAs2(str(html_path), FileFormat=10)
            
            print("✓ HTML file created")
            
            self.doc.Close(SaveChanges=False)
            self.doc = None
            
            time.sleep(1)
            
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
    window.addEventListener('DOMContentLoaded', function() {
        var content = document.body.innerHTML;
        
        content = content.replace(/MATHSTARTINLINE([\\s\\S]*?)MATHENDINLINE/g, function(match, latex) {
            return '<span class="inlineMath">' + latex + '</span>';
        });
        
        content = content.replace(/MATHSTARTDISPLAY([\\s\\S]*?)MATHENDDISPLAY/g, function(match, latex) {
            return '<div class="Math_box">' + latex + '</div>';
        });
        
        document.body.innerHTML = content;
        
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
    .inlineMath { display: inline; margin: 0 2px; color: #cc0000; }
    .Math_box { display: block; margin: 15px auto; text-align: center; color: #008800; font-size: 1.1em; }
    body { direction: rtl; text-align: right; font-family: Arial, Tahoma, sans-serif; }
    </style>
</head>
<body lang="AR-SA" dir="rtl">
""" + body_content + """
</body>
</html>"""
            
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
        """Main entry point - process document with VML awareness"""
        
        docx_path = Path(docx_path).absolute()
        
        if not output_path:
            output_path = docx_path.parent / f"{docx_path.stem}_processed.docx"
        else:
            output_path = Path(output_path).absolute()
        
        if output_path == docx_path:
            output_path = docx_path.parent / f"{docx_path.stem}_processed_safe.docx"
        
        print(f"\n{'='*60}")
        print(f"WORD COM EQUATION REPLACER (VML-Aware)")
        print(f"{'='*60}")
        print(f"📄 Input: {docx_path}")
        print(f"📄 Output: {output_path}")
        print(f"{'='*60}\n")
        
        # Step 1: Extract and analyze equations from ZIP
        self.latex_equations, xml_locations = self._extract_and_analyze_equations(docx_path)
        
        if not self.latex_equations:
            print("⚠ No equations found")
            return None
        
        try:
            # Open Word
            print("\nStarting Word application...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = False
            self.word.ScreenUpdating = False
            
            # Open document
            print("Opening document...")
            self.doc = self.word.Documents.Open(str(docx_path))
            print("✓ Document opened")
            
            # Check for tracked changes
            print("\nChecking for tracked changes...")
            has_tracked_changes = False
            try:
                # Check if track changes is enabled or if there are existing revisions
                if self.doc.TrackRevisions:
                    has_tracked_changes = True
                    print("❌ Track Changes is ENABLED in this document")
                elif self.doc.Revisions.Count > 0:
                    has_tracked_changes = True
                    print(f"❌ Document has {self.doc.Revisions.Count} tracked changes")

                if has_tracked_changes:
                    error_msg = f"Document '{docx_path.name}' has tracked changes and cannot be processed. Please accept all changes and disable tracking before processing."
                    print(f"\n{error_msg}")
                    return {
                        'error': error_msg,
                        'has_tracked_changes': True,
                        'file_name': docx_path.name
                    }
                else:
                    print("✓ No tracked changes detected")
            except Exception as e:
                print(f"⚠ Warning: Could not check tracked changes: {e}")
            
            print(f"\nDocument OMaths count: {self.doc.OMaths.Count}")
            
            # Step 2: Collect accessible equations via COM
            com_equations = self._collect_accessible_equations()
            
            # Step 3: Create mapping
            mapping = self._map_equations_smart(com_equations, xml_locations)
            
            # Step 4: Replace equations with proper mapping
            equation_count = self._replace_equations_with_mapping(com_equations, mapping)
            
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
            print(f"\n⚠️ Note: {len(self.latex_equations) - len(com_equations)} VML textbox equations remain unchanged")
            print(f"{'='*60}\n")
            
            return {
                'word_path': output_path,
                'html_path': html_path,
                'equations_replaced': equation_count,
                'equations_inaccessible': len(self.latex_equations) - len(com_equations)
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
    
    print("Starting Word COM Equation Replacer (VML-Aware)...")
    converter = WordCOMEquationReplacer()
    
    try:
        result = converter.process_document(test_file)
        if result:
            print(f"\n✅ Processing complete!")
            print(f"📄 Word: {result['word_path']}")
            print(f"🌐 HTML: {result['html_path']}")
            print(f"📊 Replaced: {result['equations_replaced']} equations")
            print(f"⚠️ Inaccessible: {result['equations_inaccessible']} VML equations")
    except Exception as e:
        print(f"\n❌ Processing failed: {e}")