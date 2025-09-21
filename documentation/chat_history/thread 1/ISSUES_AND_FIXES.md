# Common Issues and Fixes

## 1. Import Error
**Error**: `ImportError: cannot import name 'WordCOMEquationReplacer'`

**Fix**: Make sure class name in `main_word_com_equation_replacer.py` is:
```python
class WordCOMEquationReplacer:  # Not WordToHTMLConverter
```

---

## 2. Dict Return Value Error
**Error**: `TypeError: argument of type 'WindowsPath' is not iterable`

**Fix**: Handle dict returns in main.py:
```python
result = replacer.process_document(file_path, output_path)
if isinstance(result, dict):
    output_file = result.get('word_path')  # or 'html_path'
else:
    output_file = result
```

---

## 3. Unicode Logging Error
**Error**: `UnicodeEncodeError` when logging Arabic text

**Fix**: Add UTF-8 encoding to log handler:
```python
logging.FileHandler('app_debug.log', encoding='utf-8')
```

---

## 4. Wrong Output Type
**Issue**: latex_equations returning HTML instead of Word

**Fix**: Explicitly select output type:
```python
# For latex_equations - Word only
if isinstance(result, dict):
    output_file = result.get('word_path')

# For word_complete - HTML
if isinstance(result, dict):
    output_file = result.get('html_path')
```

---

## 5. Missing Equations
**Issue**: Word COM finds only 47 of 144 equations

**Fix**: Use ZIP approach for complete equation detection:
```python
USE_ZIP_APPROACH = True  # Finds all equations
```

---

## 6. Equation Markers Pattern
**Issue**: Confusion about marker format

**Correct Pattern**:
```python
# Your original requirement:
inline:  @@(\\(latex_text\\))@@
display: @@[\\[latex_text\\]]@@

# These markers allow independent search later
```

---

## Quick Debug Commands

```bash
# Check which approach is being used
grep "USE_ZIP_APPROACH" main.py

# Check class names
grep "^class" doc_processor/main_word_com_equation_replacer.py

# Check import statements
grep "from doc_processor" main.py
```