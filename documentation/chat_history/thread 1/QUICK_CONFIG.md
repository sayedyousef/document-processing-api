# Quick Configuration Guide

## Scenario 1: Windows with Full Equation Support

```python
# main.py (top of file)
USE_ZIP_APPROACH = False  # Use Word COM
```

**Result**: Uses Word automation, finds ~47 of 144 equations

---

## Scenario 2: Cross-Platform with Complete Equation Detection

```python
# main.py (top of file)
USE_ZIP_APPROACH = True  # Use ZIP approach
```

**Result**: Works on any OS, finds ALL 144 equations

---

## Scenario 3: Mixed Approach (Best of Both)

```python
# For word_complete processor only
elif processor_type == "word_complete":
    # Step 1: Use ZIP to find all equations
    from doc_processor.zip_equation_replacer import ZipEquationReplacer
    zip_replacer = ZipEquationReplacer()
    temp_docx = output_dir / f"{Path(file_path).stem}_temp.docx"
    zip_replacer.process_document(file_path, temp_docx)

    # Step 2: Use Word COM for HTML conversion
    from doc_processor.main_word_com_equation_replacer import WordCOMEquationReplacer
    word_converter = WordCOMEquationReplacer()
    result = word_converter.process_document(temp_docx, output_path)
```

**Result**: Finds ALL equations AND uses Word's HTML export

---

## Output Selection

### Get Word Document Only
```python
if isinstance(result, dict):
    output_file = result.get('word_path')
```

### Get HTML Only
```python
if isinstance(result, dict):
    output_file = result.get('html_path')
    if not output_file:
        output_file = result.get('word_path')  # Fallback
```

---

## Testing Commands

```bash
# Test equation processing (Word output)
curl -X POST http://localhost:8000/api/process \
  -F "files=@test.docx" \
  -F "processor_type=latex_equations"

# Test complete processing (HTML output)
curl -X POST http://localhost:8000/api/process \
  -F "files=@test.docx" \
  -F "processor_type=word_complete"
```