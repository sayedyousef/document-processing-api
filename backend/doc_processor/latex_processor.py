# backend/latex_processor.py
# This is YOUR EXISTING CODE - no changes needed

import os
import zipfile
import re
from lxml import etree
from omml_2_latex import DirectOmmlToLatex


def save_results(results, output_file='equations_fixed_1.tex'):
    """Save to LaTeX file - YOUR EXISTING CODE"""
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write("\\documentclass{article}\n")
        f.write("\\usepackage{amsmath}\n")
        f.write("\\usepackage{amssymb}\n")
        f.write("\\usepackage{amsfonts}\n")
        f.write("\\begin{document}\n\n")
        
        for eq in results:
            latex = eq['latex']
            
            f.write(f"% Equation {eq['index']}\n")
            f.write("\\begin{equation}\n")
            f.write(f"  {latex}\n")
            f.write("\\end{equation}\n\n")

        f.write("\\end{document}\n")


def process_word_document(docx_path):
    """Extract and convert equations - YOUR EXISTING CODE"""
    if not os.path.exists(docx_path):
        print(f"File not found: {docx_path}")
        return []
    
    parser = DirectOmmlToLatex()
    results = []
    
    with zipfile.ZipFile(docx_path, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read()
            root = etree.fromstring(content)
            
            ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
            equations = root.xpath('//m:oMath', namespaces=ns)
            
            print(f"Found {len(equations)} equations\n")
            
            for i, eq in enumerate(equations, 1):
                texts = eq.xpath('.//m:t/text()', namespaces=ns)
                text = ''.join(texts)
                
                latex = parser.parse(eq)
                results.append({
                    'index': i,
                    'text': text,
                    'latex': latex
                })
                
                print(f"Equation {i}: {latex[:50]}...")
    
    return results