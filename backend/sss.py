# ============= DIAGNOSTIC VERSION - UNDERSTAND THE 144 EQUATIONS =============
"""
This version logs detailed information about WHERE each equation is located in the XML
to understand why Word COM can't find them all
"""

import win32com.client
from pathlib import Path
import pythoncom
import zipfile
from lxml import etree
import traceback

class WordCOMEquationDiagnostic:
    """Diagnostic version to understand equation locations"""

    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        
        try:
            from doc_processor.omml_2_latex import DirectOmmlToLatex
        except ImportError:
            from doc_processor.omml_2_latex import DirectOmmlToLatex
        
        self.omml_parser = DirectOmmlToLatex()
        self.latex_equations = []

    def _analyze_equation_locations_in_xml(self, docx_path):
        """Analyze WHERE each equation is located in the XML structure"""
        
        print(f"\n{'='*60}")
        print("DETAILED XML ANALYSIS - Understanding all 144 equations")
        print(f"{'='*60}\n")
        
        results = []
        equation_contexts = []
        
        try:
            with zipfile.ZipFile(docx_path, 'r') as z:
                with z.open('word/document.xml') as f:
                    content = f.read()
                    root = etree.fromstring(content)
                    
                    ns = {
                        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
                        'v': 'urn:schemas-microsoft-com:vml',
                        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006'
                    }
                    
                    equations = root.xpath('//m:oMath', namespaces=ns)
                    
                    print(f"Found {len(equations)} total equations in document.xml\n")
                    
                    for i, eq in enumerate(equations, 1):
                        # Get equation text
                        texts = eq.xpath('.//m:t/text()', namespaces=ns)
                        text = ''.join(texts)
                        latex = self.omml_parser.parse(eq)
                        
                        # Analyze the parent structure
                        parent_chain = []
                        current = eq
                        for _ in range(10):  # Look up 10 levels
                            parent = current.getparent()
                            if parent is None:
                                break
                            tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                            parent_chain.append(tag)
                            current = parent
                        
                        # Check if in special structures
                        in_textbox = 'txbxContent' in parent_chain
                        in_shape = 'pict' in parent_chain or 'shape' in parent_chain
                        in_drawing = 'drawing' in parent_chain
                        in_alternateContent = 'AlternateContent' in parent_chain
                        in_fallback = 'Fallback' in parent_chain
                        in_table = 'tc' in parent_chain  # Table cell
                        in_header = 'hdr' in parent_chain
                        in_footer = 'ftr' in parent_chain
                        
                        # Determine location type
                        location = 'main'
                        if in_textbox:
                            location = 'textbox'
                        elif in_shape:
                            location = 'shape'
                        elif in_drawing:
                            location = 'drawing'
                        elif in_alternateContent:
                            location = 'alternateContent'
                        elif in_fallback:
                            location = 'fallback'
                        elif in_table:
                            location = 'table'
                        elif in_header:
                            location = 'header'
                        elif in_footer:
                            location = 'footer'
                        
                        equation_contexts.append({
                            'index': i,
                            'location': location,
                            'parent_chain': ' -> '.join(parent_chain[:5]),
                            'in_textbox': in_textbox,
                            'in_shape': in_shape,
                            'in_alternateContent': in_alternateContent,
                            'text_preview': text[:30] if text else '[empty]'
                        })
                        
                        results.append({
                            'index': i,
                            'text': text,
                            'latex': latex,
                            'location': location
                        })
                        
                        # Log details for interesting cases
                        if location != 'main':
                            print(f"  Equation {i}: Location={location}")
                            print(f"    Parent chain: {' -> '.join(parent_chain[:5])}")
                            if i <= 3 or i > 141:
                                print(f"    LaTeX: {latex[:50]}...")
            
            # Summary statistics
            print(f"\n{'='*40}")
            print("EQUATION LOCATION SUMMARY:")
            print(f"{'='*40}")
            
            location_counts = {}
            for ctx in equation_contexts:
                loc = ctx['location']
                location_counts[loc] = location_counts.get(loc, 0) + 1
            
            for loc, count in sorted(location_counts.items()):
                print(f"  {loc}: {count} equations")
            
            # Check for AlternateContent
            alternate_count = sum(1 for ctx in equation_contexts if ctx['in_alternateContent'])
            if alternate_count > 0:
                print(f"\n⚠️ IMPORTANT: {alternate_count} equations are in AlternateContent!")
                print("  These may not be accessible via Word COM!")
            
            # Check for textboxes/shapes
            textbox_count = sum(1 for ctx in equation_contexts if ctx['in_textbox'])
            shape_count = sum(1 for ctx in equation_contexts if ctx['in_shape'])
            if textbox_count > 0:
                print(f"\n📦 {textbox_count} equations are in TextBoxes")
            if shape_count > 0:
                print(f"\n🔷 {shape_count} equations are in Shapes/Pictures")
            
            print(f"\n✓ Analysis complete: {len(results)} equations analyzed")
            return results, equation_contexts
            
        except Exception as e:
            print(f"❌ Error analyzing XML: {e}")
            traceback.print_exc()
            return [], []

    def _find_equations_with_detailed_logging(self):
        """Find equations with detailed logging of what's found where"""
        
        print(f"\n{'='*60}")
        print("WORD COM SEARCH - Detailed logging")
        print(f"{'='*60}\n")
        
        all_equations = []
        seen_positions = set()
        
        # Method 1: Main document
        print("Searching main document...")
        try:
            main_count = self.doc.OMaths.Count
            print(f"  doc.OMaths.Count = {main_count}")
            
            for i in range(1, main_count + 1):
                eq = self.doc.OMaths.Item(i)
                pos = eq.Range.Start
                eq_text = eq.Range.Text[:30] if eq.Range.Text else '[empty]'
                
                if pos not in seen_positions:
                    seen_positions.add(pos)
                    all_equations.append({
                        'object': eq,
                        'position': pos,
                        'location': 'main',
                        'text_preview': eq_text
                    })
                    
                    if i <= 3 or i > main_count - 3:
                        print(f"    Equation {i}: pos={pos}, text={eq_text}")
        except Exception as e:
            print(f"  Error: {e}")
        
        print(f"  Found {len(all_equations)} in main\n")
        
        # Method 2: Check Shapes collection
        print("Searching Shapes collection...")
        shape_equations = 0
        try:
            print(f"  Total shapes: {self.doc.Shapes.Count}")
            
            for s_idx in range(1, self.doc.Shapes.Count + 1):
                shape = self.doc.Shapes.Item(s_idx)
                
                # Check if shape has TextFrame
                if hasattr(shape, 'TextFrame'):
                    if shape.TextFrame.HasText:
                        tf_range = shape.TextFrame.TextRange
                        if tf_range.OMaths.Count > 0:
                            print(f"    Shape {s_idx}: Has {tf_range.OMaths.Count} equations")
                            
                            for i in range(1, tf_range.OMaths.Count + 1):
                                eq = tf_range.OMaths.Item(i)
                                pos = eq.Range.Start
                                
                                if pos not in seen_positions:
                                    seen_positions.add(pos)
                                    all_equations.append({
                                        'object': eq,
                                        'position': pos,
                                        'location': f'shape_{s_idx}',
                                        'text_preview': eq.Range.Text[:30]
                                    })
                                    shape_equations += 1
                
                # Check AlternateShapeNodes
                if hasattr(shape, 'AlternateText'):
                    if 'equation' in shape.AlternateText.lower():
                        print(f"    Shape {s_idx}: Has equation in AlternateText")
        
        except Exception as e:
            print(f"  Error checking shapes: {e}")
        
        print(f"  Found {shape_equations} equations in shapes\n")
        
        # Method 3: Check InlineShapes
        print("Searching InlineShapes...")
        inline_equations = 0
        try:
            print(f"  Total inline shapes: {self.doc.InlineShapes.Count}")
            
            for i in range(1, self.doc.InlineShapes.Count + 1):
                shape = self.doc.InlineShapes.Item(i)
                
                # Check if it's an equation (Type 12)
                if shape.Type == 12:  # wdInlineShapeEquation
                    print(f"    InlineShape {i}: Is equation type")
                    inline_equations += 1
                    
                # Check for embedded objects
                if shape.Type == 1:  # wdInlineShapeEmbeddedOLEObject
                    if hasattr(shape, 'OLEFormat'):
                        print(f"    InlineShape {i}: Embedded OLE object")
        
        except Exception as e:
            print(f"  Error checking inline shapes: {e}")
        
        print(f"  Found {inline_equations} equation-type inline shapes\n")
        
        # Method 4: Story ranges with details
        print("Searching all story ranges...")
        story_types = {
            1: "MainText",
            2: "Footnotes",
            3: "Endnotes",
            4: "Comments",
            5: "TextFrame",
            6: "EvenPagesHeader",
            7: "PrimaryHeader",
            8: "EvenPagesFooter",
            9: "PrimaryFooter",
            10: "FirstPageHeader",
            11: "FirstPageFooter"
        }
        
        for story_type, name in story_types.items():
            try:
                story = self.doc.StoryRanges(story_type)
                story_count = 0
                
                while story:
                    if story.OMaths.Count > 0:
                        print(f"  {name}: {story.OMaths.Count} equations")
                        
                        for i in range(1, story.OMaths.Count + 1):
                            eq = story.OMaths.Item(i)
                            pos = eq.Range.Start
                            
                            if pos not in seen_positions:
                                seen_positions.add(pos)
                                all_equations.append({
                                    'object': eq,
                                    'position': pos,
                                    'location': name,
                                    'text_preview': eq.Range.Text[:30]
                                })
                                story_count += 1
                    
                    story = story.NextStoryRange
                
                if story_count > 0:
                    print(f"    Total in {name}: {story_count}")
                    
            except:
                pass
        
        print(f"\n✓ Total equations found by Word COM: {len(all_equations)}")
        
        return all_equations

    def diagnose_document(self, docx_path):
        """Run diagnostic analysis"""
        
        docx_path = Path(docx_path).absolute()
        
        print(f"\n{'='*60}")
        print("EQUATION DIAGNOSTIC ANALYSIS")
        print(f"{'='*60}")
        print(f"📄 Document: {docx_path.name}")
        print(f"{'='*60}\n")
        
        # Step 1: Analyze XML structure
        xml_equations, equation_contexts = self._analyze_equation_locations_in_xml(docx_path)
        
        if not xml_equations:
            print("No equations found in XML")
            return
        
        try:
            # Step 2: Open Word
            print("\nOpening document in Word...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = False
            
            self.doc = self.word.Documents.Open(str(docx_path))
            print("✓ Document opened")
            
            # Accept changes
            try:
                self.doc.AcceptAllRevisions()
            except:
                pass
            
            # Step 3: Find equations with detailed logging
            word_equations = self._find_equations_with_detailed_logging()
            
            # Step 4: Analysis
            print(f"\n{'='*60}")
            print("DIAGNOSTIC RESULTS:")
            print(f"{'='*60}")
            print(f"📊 XML equations: {len(xml_equations)}")
            print(f"📊 Word COM equations: {len(word_equations)}")
            print(f"📊 Missing equations: {len(xml_equations) - len(word_equations)}")
            
            # Find which types are missing
            xml_locations = {}
            for eq in xml_equations:
                loc = eq['location']
                xml_locations[loc] = xml_locations.get(loc, 0) + 1
            
            word_locations = {}
            for eq in word_equations:
                loc = eq['location']
                word_locations[loc] = word_locations.get(loc, 0) + 1
            
            print(f"\n📍 Equations by location in XML:")
            for loc, count in xml_locations.items():
                word_count = word_locations.get(loc, 0)
                if word_count < count:
                    print(f"  {loc}: {count} in XML, {word_count} in Word ⚠️ Missing {count - word_count}")
                else:
                    print(f"  {loc}: {count} in XML, {word_count} in Word ✓")
            
            # Check for AlternateContent
            alternate_equations = [ctx for ctx in equation_contexts if ctx['in_alternateContent']]
            if alternate_equations:
                print(f"\n⚠️ CRITICAL: {len(alternate_equations)} equations are in AlternateContent!")
                print("These equations exist in the XML but may not be accessible via Word COM")
                print("They might be compatibility fallbacks or hidden content")
                
                # Show some examples
                print("\nExamples of AlternateContent equations:")
                for ctx in alternate_equations[:5]:
                    print(f"  Equation {ctx['index']}: {ctx['parent_chain']}")
            
            return xml_equations, word_equations
            
        except Exception as e:
            print(f"❌ Error: {e}")
            traceback.print_exc()
            
        finally:
            if self.doc:
                self.doc.Close()
            if self.word:
                self.word.Quit()
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    test_file = r"C:\Users\elsayedyousef\Downloads\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx"
    
    print("Starting Equation Diagnostic Analysis...")
    diagnostic = WordCOMEquationDiagnostic()
    
    try:
        xml_eqs, word_eqs = diagnostic.diagnose_document(test_file)
        print(f"\n✅ Diagnostic complete")
    except Exception as e:
        print(f"\n❌ Diagnostic failed: {e}")