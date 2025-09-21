# ============= QUICK EQUATION DIAGNOSTIC =============
"""Fast diagnostic to identify missing equations without hanging"""

import zipfile
from pathlib import Path
from lxml import etree
import json

class QuickEquationDiagnostic:
    """Fast diagnostic using only ZIP analysis"""
    
    def diagnose_document(self, docx_path):
        """Analyze document structure from ZIP only"""
        
        docx_path = Path(docx_path).absolute()
        print(f"\n{'='*70}")
        print(f"QUICK EQUATION DIAGNOSTIC")
        print(f"Document: {docx_path.name}")
        print(f"{'='*70}\n")
        
        equations = []
        
        with zipfile.ZipFile(docx_path, 'r') as z:
            with z.open('word/document.xml') as f:
                content = f.read()
                root = etree.fromstring(content)
                
                ns = {
                    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }
                
                all_omaths = root.xpath('//m:oMath', namespaces=ns)
                
                for i, omath in enumerate(all_omaths, 1):
                    # Get equation text
                    texts = omath.xpath('.//m:t/text()', namespaces=ns)
                    text = ''.join(texts)
                    
                    # Check if in table
                    in_table = False
                    current = omath
                    for _ in range(10):  # Check up to 10 levels up
                        if current is None:
                            break
                        if current.tag.endswith('tc'):  # tc = table cell
                            in_table = True
                            break
                        current = current.getparent()
                    
                    equations.append({
                        'index': i,
                        'text': text,
                        'in_table': in_table
                    })
        
        print(f"Total equations found: {len(equations)}")
        
        # Analyze structure
        table_equations = [eq for eq in equations if eq['in_table']]
        non_table_equations = [eq for eq in equations if not eq['in_table']]
        
        print(f"\nEquation distribution:")
        print(f"  In tables:     {len(table_equations)}")
        print(f"  Not in tables: {len(non_table_equations)}")
        
        # Find where table equations start
        first_table_eq = None
        last_non_table_eq = None
        
        for i, eq in enumerate(equations):
            if eq['in_table'] and first_table_eq is None:
                first_table_eq = i + 1
                if i > 0:
                    last_non_table_eq = i
        
        if first_table_eq:
            print(f"\n⚠ KEY FINDING:")
            print(f"  Equations 1-{last_non_table_eq}: NOT in tables")
            print(f"  Equations {first_table_eq}-{len(equations)}: IN TABLES")
            print(f"\n  Word COM likely stops at equation #{last_non_table_eq} (just before tables)")
        
        # Show equation #70 and #71
        if len(equations) > 70:
            print(f"\n📍 Equation #70 (last Word COM sees):")
            eq70 = equations[69]
            print(f"  Text: {eq70['text']}")
            print(f"  In table: {eq70['in_table']}")
            
            print(f"\n📍 Equation #71 (first missing):")
            eq71 = equations[70]
            print(f"  Text: {eq71['text']}")
            print(f"  In table: {eq71['in_table']}")
            
            if not eq70['in_table'] and eq71['in_table']:
                print(f"\n✅ CONFIRMED: Word COM stops when equations enter a table!")
        
        # Show patterns
        print(f"\n📊 Equation locations by groups of 10:")
        for i in range(0, len(equations), 10):
            group = equations[i:i+10]
            in_table_count = sum(1 for eq in group if eq['in_table'])
            print(f"  Equations {i+1}-{min(i+10, len(equations))}: {in_table_count}/10 in tables")
        
        # Save report
        report_path = docx_path.parent / f"{docx_path.stem}_quick_diagnostic.json"
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump({
                'total': len(equations),
                'in_tables': len(table_equations),
                'not_in_tables': len(non_table_equations),
                'equations': equations
            }, f, indent=2, ensure_ascii=False)
        
        print(f"\n📋 Report saved to: {report_path}")
        
        return equations


if __name__ == "__main__":
    import sys
    
    # Get file from command line or use default
    if len(sys.argv) > 1:
        test_file = sys.argv[1]
    else:
        # Use a default path - change this to your file
        test_file = r"C:\Users\elsayedyousef\Downloads\الدالة واحد لواحد (جاهزة للنشر) - Copy.docx"
    
    print("Starting Quick Equation Diagnostic...")
    diagnostic = QuickEquationDiagnostic()
    diagnostic.diagnose_document(test_file)