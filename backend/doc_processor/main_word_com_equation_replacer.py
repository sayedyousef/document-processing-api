# ============= word_com_equation_replacer.py - HYBRID APPROACH =============
"""Use ZIP extraction for conversion, Word COM for replacement"""
import sys
import os
import win32com.client
from pathlib import Path
import pythoncom
import json
import zipfile
from lxml import etree
import traceback

# Add parent directory to path for imports
sys.path.append(str(Path(__file__).parent.parent))
#from core.logger import setup_logging
from core.logger import setup_logger  # <- Direct import from same directory

from omml_2_latex import DirectOmmlToLatex

logger = setup_logger("word_com_replacer")

class WordCOMEquationReplacer:
    """Hybrid approach: ZIP extraction for conversion, Word COM for replacement"""

    def __init__(self):
        pythoncom.CoInitialize()
        self.word = None
        self.doc = None
        self.omml_parser = DirectOmmlToLatex()
        self.latex_equations = []  # Will store converted equations

    def _extract_and_convert_equations(self, docx_path):
        """Extract equations from ZIP and convert to LaTeX using your method"""
        
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
                        # Extract text for reference
                        texts = eq.xpath('.//m:t/text()', namespaces=ns)
                        text = ''.join(texts)
                        
                        # Convert to LaTeX using your parser
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

    def process_document(self, docx_path, output_path=None):
        """Main entry point - Extract equations, then replace using Word COM"""
        
        docx_path = Path(docx_path).absolute()
        
        if not output_path:
            output_path = docx_path.parent / f"{docx_path.stem}_equations_text.docx"
        else:
            output_path = Path(output_path).absolute()


        if output_path == docx_path:
            print(f"❌ ERROR: Output path same as input!")
            print(f"❌ Would overwrite original document!")
            # Force a different name
            output_path = docx_path.parent / f"{docx_path.stem}_equations_text_safe.docx"
            print(f"✓ Changed output to: {output_path}")
        
        # SAFETY CHECK: If output already exists, add number
        if output_path.exists():
            counter = 1
            while True:
                new_output = output_path.parent / f"{output_path.stem}_{counter}{output_path.suffix}"
                if not new_output.exists():
                    output_path = new_output
                    break
                counter += 1
            print(f"✓ Output file exists, using: {output_path}")

        print(f"\n{'='*60}")
        print(f"📁 Input: {docx_path}")
        print(f"📁 Output: {output_path}")
        print(f"{'='*60}\n")
        
        # STEP 1: Extract and convert equations using ZIP method
        self.latex_equations = self._extract_and_convert_equations(docx_path)
        
        if not self.latex_equations:
            print("⚠ No equations extracted, aborting...")
            return None
        
        try:
            # STEP 2: Open document with Word COM for replacement
            print(f"\n{'='*40}")
            print("STEP 2: Replacing equations using Word COM")
            print(f"{'='*40}\n")
            
            print("Starting Word application...")
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = False  # Don't show alerts
            self.word.ScreenUpdating = False  # Don't update screen

            print(f"Opening document...")
            self.doc = self.word.Documents.Open(str(docx_path))
            print(f"✓ Document opened successfully")
            
            # Replace equations
            equation_count = self._replace_equations()
            
            # Save modified document
            print(f"\nSaving modified document...")
            self.doc.SaveAs2(str(output_path))
            print(f"✓ Document saved")
            
            print(f"\n{'='*60}")
            print(f"✅ SUCCESS!")
            print(f"📄 Output file: {output_path}")
            print(f"📊 Equations replaced: {equation_count}")
            print(f"{'='*60}\n")
            
            return output_path
            
        except Exception as e:
            print(f"\n❌ ERROR in process_document: {e}")
            traceback.print_exc()
            raise
            
        finally:
                self._cleanup()

    def _replace_equations(self):
        """Replace equations: Delete equation, Insert LaTeX text with markers"""
        
        equations_replaced = 0
        equations_data = []
        
        # Get initial count
        initial_count = self.doc.OMaths.Count
        
        print(f"📐 Word COM found {initial_count} equations")
        print(f"📝 We have {len(self.latex_equations)} converted LaTeX equations")
        
        if initial_count != len(self.latex_equations):
            print(f"⚠ WARNING: Equation count mismatch!")
            print(f"  Word COM: {initial_count}")
            print(f"  LaTeX array: {len(self.latex_equations)}")
        
        print("-" * 40)
        
        # Track which LaTeX equation we're using from our array
        latex_index = 0
        
        # Process equations - always work with Item(1) since we're deleting
        while self.doc.OMaths.Count > 0 and latex_index < len(self.latex_equations):
            remaining = self.doc.OMaths.Count
            print(f"\nProcessing equation {latex_index + 1} (remaining in Word: {remaining})...")
            
            try:
                # STEP 1: Get the first equation (always Item(1) since we delete)
                omath = self.doc.OMaths.Item(1)
                print(f"  Got OMath object")
                
                # STEP 2: Get replacement LaTeX from our array
                latex_data = self.latex_equations[latex_index]
                latex_text = latex_data['latex']
                print(f"  LaTeX from array[{latex_index}]: {latex_text[:50]}..." if len(latex_text) > 50 else f"  LaTeX from array[{latex_index}]: {latex_text}")
                
                # Clean the LaTeX text
                if latex_text:
                    latex_text = latex_text.strip()
                if not latex_text:
                    latex_text = f"[EQUATION_{latex_index + 1}_EMPTY]"
                
                # STEP 3: Get the range where equation is (before deleting)
                eq_range = omath.Range
                
                # STEP 4: DELETE the equation
                print(f"  Deleting equation...")
                try:
                    eq_range.Select()
                    self.word.Selection.Delete()
                    print(f"  ✓ Equation deleted")
                except Exception as del_error:
                    print(f"  ❌ Delete failed: {del_error}")
                    latex_index += 1
                    continue
                
                # STEP 5: INSERT the LaTeX text WITH MARKERS
                print(f"  Inserting LaTeX text with MathJax delimiters...")
                equation_id = f"eq_{latex_index + 1}"
                bookmark_name = equation_id  # Define bookmark_name

                try:
                    # Determine if equation is inline or display based on length
                    is_inline = len(latex_text) < 30  # Simple heuristic
                    
                    # Use MathJax delimiters that will be recognized automatically
                    if is_inline:
                        marked_text = f"\\({latex_text}\\)"  # Inline equation
                        print(f"  Using inline format (length: {len(latex_text)})")
                    else:
                        marked_text = f"\\[{latex_text}\\]"  # Display equation
                        print(f"  Using display format (length: {len(latex_text)})")
                    
                    # Insert the marked text with spaces
                    self.word.Selection.TypeText(f" {marked_text} ")
                    
                    # Optional: Also add Word bookmark (won't appear in HTML but useful in Word)
                    try:
                        # Move back to select what we just typed
                        self.word.Selection.MoveLeft(Count=len(marked_text) + 2)
                        self.word.Selection.MoveRight(Count=len(marked_text) + 2, Extend=True)
                        
                        # Add bookmark
                        self.doc.Bookmarks.Add(bookmark_name, self.word.Selection.Range)
                        print(f"  ✓ Bookmark added: {bookmark_name}")
                    except:
                        pass  # Bookmark is optional
                    
                    print(f"  ✓ Text inserted with MathJax delimiters")
                    equations_replaced += 1
                    
                except Exception as insert_error:
                    print(f"  ❌ Insert failed: {insert_error}")

                # Store equation data
                equations_data.append({
                    'index': latex_index + 1,
                    'latex': latex_text,
                    'bookmark': bookmark_name,
                    'status': 'replaced',
                    'type': 'inline' if is_inline else 'display'  # Track the type
                })
                
                # Move to next LaTeX in our array
                latex_index += 1
                
            except Exception as e:
                print(f"❌ Critical error for equation {latex_index + 1}: {e}")
                traceback.print_exc()
                
                # Try to skip this equation
                try:
                    self.doc.OMaths.Item(1).Range.Delete()
                    print(f"  Deleted problematic equation")
                except:
                    print(f"  Could not delete, stopping...")
                    break
                
                latex_index += 1
        
        # Check if we have leftover equations
        if self.doc.OMaths.Count > 0:
            print(f"\n⚠ WARNING: {self.doc.OMaths.Count} equations remain in Word!")
        
        if latex_index < len(self.latex_equations):
            print(f"⚠ WARNING: {len(self.latex_equations) - latex_index} LaTeX equations unused!")
        
        # Save equation data
        self._save_equation_data(equations_data)
        
        # Summary
        print(f"\n{'='*40}")
        print(f"SUMMARY:")
        print(f"  Initial equations in Word: {initial_count}")
        print(f"  LaTeX equations in array: {len(self.latex_equations)}")
        print(f"  Successfully replaced: {equations_replaced}")
        print(f"  Equations processed: {latex_index}")
        print(f"{'='*40}")
        
        return equations_replaced


    def _save_equation_data(self, equations_data):
        """Save equation data to JSON file"""
        if self.doc and self.doc.FullName:
            json_path = Path(self.doc.FullName).parent / f"{Path(self.doc.Name).stem}_equations.json"
            
            # Include both Word replacements and original LaTeX conversions
            full_data = {
                'source_document': str(self.doc.FullName),
                'latex_conversions': self.latex_equations,
                'word_replacements': equations_data,
                'total_equations': len(equations_data)
            }
            
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(full_data, f, indent=2, ensure_ascii=False)
            
            print(f"\n📋 Equation data saved to: {json_path}")

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


# ============= Main execution =============
if __name__ == "__main__":
    test_file = r"D:\Work 3 (20-Oct-24)\2 Side projects May 25\Encyclopedia\articles\مقالات بعد الاخراج\test\الدالة واحد لواحد (جاهزة للنشر).docx"
    
    print("Starting Hybrid Equation Replacer...")
    print("Using ZIP extraction for conversion, Word COM for replacement")
    
    processor = WordCOMEquationReplacer()
    
    try:
        output = processor.process_document(test_file)
        if output:
            print(f"\n✅ Processing complete!")
            print(f"📄 Output file: {output}")
    except Exception as e:
        print(f"\n❌ Processing failed: {e}")
        traceback.print_exc()