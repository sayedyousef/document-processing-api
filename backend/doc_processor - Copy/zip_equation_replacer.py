# ============= ZIP EQUATION REPLACER WITH TRACK CHANGES HANDLING =============
"""
ZIP-based equation replacer that handles Track Changes
No Word COM needed - cleans tracked changes directly in XML
"""
import zipfile
import os
import shutil
from pathlib import Path
from lxml import etree
import traceback
from .omml_2_latex import DirectOmmlToLatex

class ZipEquationReplacer:
    """ZIP-based equation replacer - handles Track Changes without Word COM"""
    
    def __init__(self):
        self.omml_parser = DirectOmmlToLatex()
        self.equations_found = []
    
    def process_document(self, docx_path, output_path=None):
        """Process document using ZIP approach - handles Track Changes automatically"""
        
        docx_path = Path(docx_path).absolute()
        
        if not output_path:
            output_path = docx_path.parent / f"{docx_path.stem}_latex_equations.docx"
        else:
            output_path = Path(output_path).absolute()
        
        # Safety: Never overwrite original
        if output_path == docx_path:
            output_path = docx_path.parent / f"{docx_path.stem}_latex_equations_safe.docx"
        
        print(f"\n{'='*60}")
        print(f"🔒 ZIP EQUATION REPLACER (Track Changes Aware)")
        print(f"📁 Input: {docx_path}")
        print(f"📁 Output: {output_path}")
        print(f"{'='*60}\n")
        
        try:
            # STEP 1: Create clean version without Track Changes
            temp_clean = docx_path.parent / f"{docx_path.stem}_clean_temp.docx"
            self.accept_all_changes_and_disable_tracking(docx_path, temp_clean)
            
            # STEP 2: Process equations on the CLEAN document
            equations = self._extract_and_convert_equations_from_zip(temp_clean)
            
            if not equations:
                print("⚠ No equations found, copying clean document")
                shutil.copy2(temp_clean, output_path)
                temp_clean.unlink()  # Clean up
                return output_path
            
            print(f"✓ Found {len(equations)} equations")
            
            # STEP 3: Replace equations in the clean document
            self._replace_equations_in_zip(temp_clean, output_path, equations)
            
            # Clean up temp file
            temp_clean.unlink()
            
            print(f"\n✅ SUCCESS! ZIP processing complete")
            print(f"📄 Output: {output_path}")
            print(f"📊 Equations processed: {len(equations)}")
            
            return output_path
            
        except Exception as e:
            print(f"❌ ERROR: {e}")
            traceback.print_exc()
            # Clean up temp file if it exists
            temp_clean = docx_path.parent / f"{docx_path.stem}_clean_temp.docx"
            if temp_clean.exists():
                temp_clean.unlink()
            shutil.copy2(docx_path, output_path)
            return output_path
    
    def accept_all_changes_and_disable_tracking(self, docx_path, output_path):
        """
        Complete solution: Accept ALL tracked changes and disable tracking
        """
        
        print("\n" + "="*60)
        print("Processing Track Changes in ZIP")
        print("="*60)
        
        with zipfile.ZipFile(docx_path, 'r') as zip_in:
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_out:
                
                for item in zip_in.infolist():
                    
                    # STEP 1: Clean document.xml - Accept all changes
                    if item.filename == 'word/document.xml':
                        content = zip_in.read(item.filename)
                        root = etree.fromstring(content)
                        
                        cleaned_root = self._accept_all_tracked_changes(root)
                        
                        modified_content = etree.tostring(
                            cleaned_root, 
                            encoding='UTF-8', 
                            xml_declaration=True,
                            pretty_print=True
                        )
                        zip_out.writestr(item, modified_content)
                        
                    # STEP 2: Modify settings.xml - Turn OFF tracking
                    elif item.filename == 'word/settings.xml':
                        content = zip_in.read(item.filename)
                        root = etree.fromstring(content)
                        
                        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                        
                        # Remove trackRevisions element (turns OFF tracking)
                        track_elem = root.find('.//w:trackRevisions', namespaces=ns)
                        if track_elem is not None:
                            parent = track_elem.getparent()
                            parent.remove(track_elem)
                            print("  ✓ Track Changes disabled")
                        
                        modified_content = etree.tostring(root, encoding='UTF-8', xml_declaration=True)
                        zip_out.writestr(item, modified_content)
                        
                    # STEP 3: Skip people.xml and revisionsView.xml (no longer needed)
                    elif item.filename in ['word/people.xml', 'word/revisionsView.xml']:
                        print(f"  Skipping {item.filename} (no longer needed)")
                        continue
                        
                    else:
                        # Copy all other files
                        zip_out.writestr(item, zip_in.read(item.filename))
        
        print("✓ All changes accepted, tracking disabled")
    
    def _accept_all_tracked_changes(self, root):
        """
        Accept all tracked changes in document.xml - COMPLETE IMPLEMENTATION
        """
        
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        changes_count = {
            'insertions': 0,
            'deletions': 0,
            'format_changes': 0,
            'moves': 0
        }
        
        # 1. Process DELETIONS first (remove them)
        for del_elem in root.xpath('//w:del', namespaces=ns):
            parent = del_elem.getparent()
            if parent is not None:
                parent.remove(del_elem)
                changes_count['deletions'] += 1
        
        # 2. Process INSERTIONS (keep content, remove wrapper)
        for ins_elem in root.xpath('//w:ins', namespaces=ns):
            parent = ins_elem.getparent()
            if parent is not None:
                # Move all children out of w:ins wrapper
                for child in list(ins_elem):
                    parent.insert(parent.index(ins_elem), child)
                parent.remove(ins_elem)
                changes_count['insertions'] += 1
        
        # 3. Process MOVES (moveFrom/moveTo)
        # Remove moveFrom (source of move)
        for move_from in root.xpath('//w:moveFrom', namespaces=ns):
            parent = move_from.getparent()
            if parent is not None:
                parent.remove(move_from)
                changes_count['moves'] += 1
        
        # Keep moveTo content (destination of move)
        for move_to in root.xpath('//w:moveTo', namespaces=ns):
            parent = move_to.getparent()
            if parent is not None:
                for child in list(move_to):
                    parent.insert(parent.index(move_to), child)
                parent.remove(move_to)
        
        # 4. Process FORMAT CHANGES (remove change tracking attributes)
        for elem in root.xpath('//*[@w:rsidR or @w:rsidDel or @w:rsidRPr or @w:rsidTr]', namespaces=ns):
            # Remove all revision tracking attributes
            attrs_to_remove = ['rsidR', 'rsidDel', 'rsidRPr', 'rsidTr', 'rsidP', 'rsidRDefault']
            for attr in attrs_to_remove:
                elem.attrib.pop(f'{{{ns["w"]}}}{attr}', None)
        
        # 5. Remove property changes
        for prop_change in root.xpath('//w:pPrChange | //w:rPrChange', namespaces=ns):
            parent = prop_change.getparent()
            if parent is not None:
                parent.remove(prop_change)
                changes_count['format_changes'] += 1
        
        print(f"\n  Changes accepted:")
        print(f"    Deletions removed: {changes_count['deletions']}")
        print(f"    Insertions accepted: {changes_count['insertions']}")
        print(f"    Moves processed: {changes_count['moves']}")
        print(f"    Format changes: {changes_count['format_changes']}")
        
        return root
    
    def _extract_and_convert_equations_from_zip(self, docx_path):
        """Extract equations from ZIP file"""
        
        print(f"\n{'='*40}")
        print("Extracting equations from ZIP")
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
    
    def _replace_equations_in_zip(self, input_path, output_path, equations):
        """Replace equations in ZIP file"""
        
        print(f"\n{'='*40}")
        print("Replacing equations in ZIP")
        print(f"{'='*40}")
        
        try:
            # Read the document.xml from the input ZIP
            with zipfile.ZipFile(input_path, 'r') as z:
                with z.open('word/document.xml') as f:
                    content = f.read()
                    root = etree.fromstring(content)
            
            # Replace equations in the XML
            self._replace_equations_in_xml(root, equations)
            
            # Create a NEW ZIP file (not append mode!)
            with zipfile.ZipFile(input_path, 'r') as zip_in:
                with zipfile.ZipFile(output_path, 'w') as zip_out:
                    # Copy all files from input ZIP
                    for item in zip_in.infolist():
                        if item.filename == 'word/document.xml':
                            # Replace document.xml with our modified version
                            modified_content = etree.tostring(root, encoding='unicode', pretty_print=True)
                            zip_out.writestr(item, modified_content.encode('utf-8'))
                        else:
                            # Copy all other files as-is
                            zip_out.writestr(item, zip_in.read(item.filename))
                
            print(f"✓ Equations replaced in ZIP successfully")
            
        except Exception as e:
            print(f"❌ Error replacing equations in ZIP: {e}")
            traceback.print_exc()
            raise


    def _replace_equations_in_xml(self, root, equations):
        """Replace equations in XML with markers for HTML processing"""
        
        ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        all_equations = root.xpath('//m:oMath', namespaces=ns)
        
        print(f"Found {len(all_equations)} equations to replace")
        
        equations_replaced = 0
        
        for i, eq_node in enumerate(all_equations):
            if i >= len(equations):
                break
            
            try:
                latex = equations[i]['latex'].strip() or f"[EQUATION_{i + 1}_EMPTY]"
                
                # Determine if inline or display
                is_inline = len(latex) < 30
                
                # Create marked text with MATHSTARTINLINE/MATHSTARTDISPLAY
                if is_inline:
                    marked_text = f' MATHSTARTINLINE\\({latex}\\)MATHENDINLINE '
                else:
                    marked_text = f' MATHSTARTDISPLAY\\[{latex}\\]MATHENDDISPLAY '
                
                # Get parent
                parent = eq_node.getparent()
                
                if parent is not None:
                    parent_tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                    
                    # Create text element with markers
                    if parent_tag == 'r':
                        # In a run - replace with text
                        t = etree.Element(f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        t.text = marked_text
                        parent.replace(eq_node, t)
                    else:
                        # In paragraph or elsewhere - create run with text
                        r = etree.Element(f'{{{ns["w"]}}}r')
                        t = etree.SubElement(r, f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        t.text = marked_text
                        parent.replace(eq_node, r)
                    
                    equations_replaced += 1
                    print(f"  Replaced equation {i+1}: {latex[:30]}...")
                    
            except Exception as e:
                print(f"Error replacing equation {i+1}: {e}")
        
        print(f"✓ Replaced {equations_replaced} equations with markers")
        return root

    def _replace_equations_in_xml_old(self, root, equations):
        """Replace equations in XML - handles all equation types"""
        
        ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        # Get all equations at once (they won't change as we process)
        all_equations = root.xpath('//m:oMath', namespaces=ns)
        
        if len(all_equations) != len(equations):
            print(f"⚠ WARNING: Found {len(all_equations)} equations but have {len(equations)} replacements")
        
        equations_replaced = 0
        
        # Process each equation
        for i, eq_node in enumerate(all_equations):
            if i >= len(equations):
                print(f"⚠ No more LaTeX equations for equation {i+1}")
                break
            
            print(f"Processing equation {i + 1}...")
            
            try:
                latex = equations[i]['latex']
                
                # Clean the LaTeX text
                if latex:
                    latex = latex.strip()
                if not latex:
                    latex = f"[EQUATION_{i + 1}_EMPTY]"
                
                # Get parent to determine context
                parent = eq_node.getparent()
                
                if parent is not None:
                    parent_tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                    
                    # CASE 1: Equation is in a run (inline equation)
                    if parent_tag == 'r':
                        # Replace with text element
                        t = etree.Element(f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        t.text = f" \\({latex}\\) "  # Inline format
                        parent.replace(eq_node, t)
                        print(f"    ✓ Replaced inline equation")
                        
                    # CASE 2: Equation is in a paragraph
                    elif parent_tag == 'p':
                        # Check if equation is the only child (block equation)
                        is_block = len([e for e in parent if e.tag.endswith('r') or e.tag.endswith('oMath')]) == 1
                        
                        # Create a run with text
                        r = etree.Element(f'{{{ns["w"]}}}r')
                        t = etree.SubElement(r, f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        
                        if is_block:
                            t.text = f"\\[{latex}\\]"  # Display format
                            print(f"    ✓ Replaced block equation")
                        else:
                            t.text = f" \\({latex}\\) "  # Inline format
                            print(f"    ✓ Replaced inline equation in paragraph")
                        
                        # Replace equation with run
                        parent.replace(eq_node, r)
                        
                    # CASE 3: Equation is elsewhere
                    else:
                        print(f"    ⚠ Equation in unexpected parent: {parent_tag}")
                        # Try to create a text element
                        t = etree.Element(f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        t.text = f" \\({latex}\\) "
                        parent.replace(eq_node, t)
                    
                    equations_replaced += 1
                    print(f"    LaTeX: {latex[:40]}..." if len(latex) > 40 else f"    LaTeX: {latex}")
                
            except Exception as e:
                print(f"❌ Error processing equation {i + 1}: {e}")
                continue
        
        print(f"\n✓ Successfully replaced {equations_replaced} equations")
        return root
    
    def _extract_and_convert_equations_from_xml(self, xml_root):
        """Extract equations from cleaned XML root"""
        
        print(f"\n{'='*40}")
        print("Extracting equations from cleaned XML")
        print(f"{'='*40}")
        
        results = []
        
        try:
            ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
            equations = xml_root.xpath('//m:oMath', namespaces=ns)
            
            print(f"Found {len(equations)} equations in cleaned XML\n")
            
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
    
    def _replace_equations_in_cleaned_xml(self, xml_root, equations):
        """Replace equations in the cleaned XML"""
        
        print(f"\n{'='*40}")
        print("Replacing equations in cleaned XML")
        print(f"{'='*40}\n")
        
        ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
              'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        # Get all equations at once (they won't change as we process)
        all_equations = xml_root.xpath('//m:oMath', namespaces=ns)
        
        if len(all_equations) != len(equations):
            print(f"⚠ WARNING: Found {len(all_equations)} equations but have {len(equations)} replacements")
        
        equations_replaced = 0
        
        # Process each equation
        for i, eq_node in enumerate(all_equations):
            if i >= len(equations):
                print(f"⚠ No more LaTeX equations for equation {i+1}")
                break
            
            print(f"Processing equation {i + 1}...")
            
            try:
                latex = equations[i]['latex']
                
                # Clean the LaTeX text
                if latex:
                    latex = latex.strip()
                if not latex:
                    latex = f"[EQUATION_{i + 1}_EMPTY]"
                
                # Get parent to determine context
                parent = eq_node.getparent()
                
                if parent is not None:
                    parent_tag = parent.tag.split('}')[-1] if '}' in parent.tag else parent.tag
                    
                    # CASE 1: Equation is in a run (inline equation)
                    if parent_tag == 'r':
                        # Replace with text element
                        t = etree.Element(f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        t.text = f" \\({latex}\\) "  # Inline format
                        parent.replace(eq_node, t)
                        print(f"    ✓ Replaced inline equation")
                        
                    # CASE 2: Equation is in a paragraph
                    elif parent_tag == 'p':
                        # Check if equation is the only child (block equation)
                        is_block = len([e for e in parent if e.tag.endswith('r') or e.tag.endswith('oMath')]) == 1
                        
                        # Create a run with text
                        r = etree.Element(f'{{{ns["w"]}}}r')
                        t = etree.SubElement(r, f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        
                        if is_block:
                            t.text = f"\\[{latex}\\]"  # Display format
                            print(f"    ✓ Replaced block equation")
                        else:
                            t.text = f" \\({latex}\\) "  # Inline format
                            print(f"    ✓ Replaced inline equation in paragraph")
                        
                        # Replace equation with run
                        parent.replace(eq_node, r)
                        
                    # CASE 3: Equation is elsewhere
                    else:
                        print(f"    ⚠ Equation in unexpected parent: {parent_tag}")
                        # Try to create a text element
                        t = etree.Element(f'{{{ns["w"]}}}t')
                        t.set(f'{{{ns["w"]}}}space', 'preserve')
                        t.text = f" \\({latex}\\) "
                        parent.replace(eq_node, t)
                    
                    equations_replaced += 1
                    print(f"    LaTeX: {latex[:40]}..." if len(latex) > 40 else f"    LaTeX: {latex}")
                
            except Exception as e:
                print(f"❌ Error processing equation {i + 1}: {e}")
                continue
        
        print(f"\n✓ Successfully replaced {equations_replaced} equations")
        return xml_root
    
    def _create_output_document(self, input_path, output_path, modified_xml):
        """Create output document with modified XML"""
        
        print(f"\nCreating output document...")
        
        # Create a new ZIP file with modified document.xml
        with zipfile.ZipFile(input_path, 'r') as zip_in:
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_DEFLATED) as zip_out:
                # Copy all files from input ZIP
                for item in zip_in.infolist():
                    if item.filename == 'word/document.xml':
                        # Write our modified XML
                        modified_content = etree.tostring(
                            modified_xml, 
                            encoding='UTF-8', 
                            xml_declaration=True,
                            pretty_print=True
                        )
                        zip_out.writestr(item, modified_content)
                    else:
                        # Copy all other files as-is
                        zip_out.writestr(item, zip_in.read(item.filename))
        
        print(f"✓ Output document created")


# Test function
if __name__ == "__main__":
    test_file = r"test_document_with_track_changes.docx"
    
    print("Starting ZIP Equation Replacer (Track Changes Aware)...")
    print("This version handles Track Changes without Word COM")
    
    processor = ZipEquationReplacer()
    
    try:
        output = processor.process_document(test_file)
        if output:
            print(f"\n✅ Processing complete!")
            print(f"📄 Output file: {output}")
    except Exception as e:
        print(f"\n❌ Processing failed: {e}")
        traceback.print_exc()