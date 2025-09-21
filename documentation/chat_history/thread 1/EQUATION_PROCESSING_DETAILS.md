# Equation Processing - Technical Details

## Core Components

### 1. omml_2_latex.py
- **Purpose**: Converts OMML (Office Math ML) to LaTeX
- **Size**: ~817 lines
- **Class**: `DirectOmmlToLatex`

### 2. Word COM Approach
**File**: `main_word_com_equation_replacer.py` (353 lines)

**How it works**:
1. Opens Word using COM automation
2. Iterates through document ranges
3. Finds equation objects
4. Converts to LaTeX using omml_2_latex
5. Replaces with HTML-wrapped LaTeX markers

**Advantages**:
- Direct Word integration
- Can save as different formats

**Limitations**:
- Windows only
- May miss some equations (found 47 of 144 in test)

### 3. ZIP Approach
**File**: `zip_equation_replacer.py` (541 lines)

**How it works**:
1. Extracts .docx as ZIP
2. Parses document.xml
3. Finds all OMML equation nodes
4. Converts to LaTeX
5. Replaces in XML
6. Repackages as .docx

**Advantages**:
- Cross-platform
- Finds ALL equations (144 of 144)
- No Word dependency

**Limitations**:
- More complex XML manipulation

## Switching Between Approaches

```python
# In main.py
USE_ZIP_APPROACH = False  # False = Word COM, True = ZIP

# The code automatically imports the right replacer:
if USE_ZIP_APPROACH:
    from doc_processor.zip_equation_replacer import ZipEquationReplacer
    replacer = ZipEquationReplacer()
else:
    from doc_processor.main_word_com_equation_replacer import WordCOMEquationReplacer
    replacer = WordCOMEquationReplacer()
```

## Equation Detection Stats

Test document: Arabic math document with 144 equations

| Approach | Equations Found | Success Rate |
|----------|----------------|--------------|
| Word COM | 47             | 32.6%        |
| ZIP      | 144            | 100%         |

## Recommendation

**Use ZIP approach** for production:
- Finds all equations reliably
- Works on any platform
- More predictable results

**Use Word COM** only if:
- Need specific Word features
- Windows-only environment
- Need Word's HTML export