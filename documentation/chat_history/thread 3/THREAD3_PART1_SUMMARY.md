# Thread 3 - Part 1: VML Textbox Access Attempts

## Main Topic
Trying to add VML textbox equation access to Word COM

## Problem
- Document has 144 equations total
- Word COM can only see 76 equations
- 68 equations are trapped in VML textboxes

## VML Structure Issue
```xml
<mc:AlternateContent>
    <mc:Choice>
        <w:drawing>  <!-- Modern format -->
    </mc:Choice>
    <mc:Fallback>
        <w:pict>
            <v:shape>  <!-- Legacy VML -->
                <v:textbox>
                    <w:txbxContent>
                        <!-- EQUATIONS TRAPPED HERE -->
                    </w:txbxContent>
                </v:textbox>
            </v:shape>
        </w:pict>
    </mc:Fallback>
</mc:AlternateContent>
```

## Attempted Solutions

### Method 6: VML Textbox Access
```python
def _collect_vml_textbox_equations(self):
    """Try to access VML textboxes"""
    print("\nMethod 6: Accessing VML textboxes...")
    vml_equations = []

    for i in range(1, self.doc.Shapes.Count + 1):
        shape = self.doc.Shapes.Item(i)
        if hasattr(shape, 'TextFrame'):
            if shape.TextFrame.HasText:
                tr = shape.TextFrame.TextRange
                if tr.OMaths.Count > 0:
                    # Process equations
```

**Result**: Found 0 VML equations - COM cannot access them

## Key Finding
**VML textboxes in Fallback sections are invisible to Word COM API**