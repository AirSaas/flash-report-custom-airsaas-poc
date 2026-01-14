# Generate PPT via python-pptx

Generate a PowerPoint presentation using python-pptx for programmatic control.

## Output Structure (per spec section 5.1)

| Slide | Content |
|-------|---------|
| 1 | **Summary** - List of all projects with mood/status |
| 2 to N | **Project slides** - One slide per project with details |
| Last | **Data Notes** - List of unfilled fields and API limitations |

## Script Commands

```bash
# Generate PPT (includes automatic template verification)
python3 scripts/generate_ppt.py

# Verify template compatibility only
python3 scripts/generate_ppt.py --verify

# Analyze template and show all shape positions
python3 scripts/generate_ppt.py --analyze
```

## Instructions

### Phase 0: Template Verification (AUTOMATIC)

The script **automatically verifies** the template before generating:

1. Checks shape positions against `EXPECTED_SHAPE_POSITIONS` constant
2. Reports warnings if shapes have moved or are missing
3. Continues generating even if some shapes don't match (with warning)

**Success output:**
```
Verifying template compatibility...
✓ Template verified OK
```

**If verification fails:**
```
⚠️  Template may have changed! Only 60% of expected shapes found.
   Run: python3 scripts/generate_ppt.py --analyze
   Then update EXPECTED_SHAPE_POSITIONS in the script.
```

**To fix template sync issues:** Run `/mapping` which will analyze the template and update `EXPECTED_SHAPE_POSITIONS` automatically.

### Phase 1: Load Data

1. **Find Latest Fetched Data**
   - Look in `data/` folder for most recent `{date}_projects.json`
   - If no data found, prompt user to run `/fetch` first

2. **Load Mapping Configuration**
   - Read `config/mapping.json`
   - Identify fields and their sources/transforms

3. **Load Template**
   - Check `templates/` for PPTX file
   - **CRITICAL:** Template is the source of truth for styling

4. **Load Reference Data**
   - Extract moods, statuses, risks from fetched data
   - Build lookup tables for label resolution

### Phase 2: Template Analysis (CRITICAL)

**Before generating, ALWAYS analyze the current template to extract:**

1. **Slide Dimensions**
   ```python
   prs = Presentation(template_path)
   width = prs.slide_width.inches
   height = prs.slide_height.inches
   ```

2. **Shape Positions and Sizes**
   ```python
   for slide_idx, slide in enumerate(prs.slides):
       print(f"Slide {slide_idx}:")
       for shape in slide.shapes:
           x = round(shape.left.inches, 2)
           y = round(shape.top.inches, 2)
           w = round(shape.width.inches, 2)
           h = round(shape.height.inches, 2)
           print(f"  {shape.name}: pos=({x}, {y}) size={w}x{h}")
           if shape.has_text_frame:
               print(f"    Text: '{shape.text[:50]}...'")
   ```

3. **Available Layouts**
   ```python
   for i, layout in enumerate(prs.slide_layouts):
       print(f"Layout {i}: {layout.name}")
   ```

4. **Images and Visual Elements**
   - Count images per slide
   - Note shape types (pictures, shapes with fills, etc.)

### Phase 3: Template Preservation Rules

**MANDATORY: When working with templates, preserve ALL visual elements:**

#### 1. Duplicate Slides via XML (not add_slide)
```python
def duplicate_slide(prs, slide_index):
    """Clone a slide preserving ALL content: images, borders, colors, fonts."""
    from copy import deepcopy
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    source_slide = prs.slides[slide_index]
    slide_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Remove default shapes from new slide
    spTree = new_slide.shapes._spTree
    shapes_to_remove = []
    for child in spTree:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp']:
            shapes_to_remove.append(child)
    for shape_el in shapes_to_remove:
        spTree.remove(shape_el)

    # Copy ALL shapes from source (includes images, styled shapes, etc.)
    source_spTree = source_slide.shapes._spTree
    for child in source_spTree:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp']:
            new_el = deepcopy(child)
            spTree.append(new_el)

    # Copy image relationships
    for rel in source_slide.part.rels.values():
        if rel.reltype == RT.IMAGE:
            new_slide.part.relate_to(rel._target, RT.IMAGE)

    return new_slide
```

#### 2. Find Shapes by Position (names change after duplication)
```python
def find_shape_by_pos(slide, target_x, target_y, tolerance=0.15):
    """Find shape by position using Euclidean distance."""
    best_match = None
    best_distance = float('inf')

    for shape in slide.shapes:
        if shape.has_text_frame:
            x = round(shape.left.inches, 1)
            y = round(shape.top.inches, 1)
            dist = ((x - target_x) ** 2 + (y - target_y) ** 2) ** 0.5
            if dist <= tolerance and dist < best_distance:
                best_match = shape
                best_distance = dist

    return best_match
```

#### 3. Delete Placeholders Entirely (not just clear text)
```python
def delete_placeholder_shapes(slide):
    """Remove placeholder shapes completely to avoid 'Double click to edit'."""
    shapes_to_delete = [s for s in slide.shapes if s.is_placeholder]
    for shape in shapes_to_delete:
        sp = shape._element
        sp.getparent().remove(sp)
```

#### 4. Handle Section Titles in Template Background
If template has visual titles (not in text shapes), add leading newline:
```python
# For shapes where visual title exists above the text area
set_shape_text(shape, "\n" + content, font_size=FONT_SIZE)
```

### Phase 4: Font Size Guidelines

Use smaller fonts for dense content:

| Element | Recommended Size |
|---------|------------------|
| Slide title | 14-20pt |
| Section headers | 10-12pt |
| Body content | 8-10pt |
| Table content | 7-8pt |
| Date/metadata | 8pt |
| Dense notes | 6-7pt |

**Always test that text fits within shape bounds.**

### Phase 5: API Field Names (CRITICAL)

Always use these exact field names from the API:

| Purpose | Correct Field | WRONG (don't use) |
|---------|---------------|-------------------|
| Budget BAC | `budget_capex_initial` | budget_capex |
| Budget Actual | `budget_capex_used` | budget_actual |
| Budget EAC | `budget_capex_landing` | budget_eac |
| Milestone done | `status == 'done'` | is_completed |
| Effort planned | `effort` | effort_planned |
| Effort used | `effort_used` | effort_actual |

### Phase 6: Generate and Save

```python
output_path = f"outputs/{datetime.now().strftime('%Y-%m-%d')}_portfolio_skill.pptx"
prs.save(output_path)
```

## Using the Script Directly

```bash
python3 scripts/generate_ppt.py
```

---

## Styling Guidelines

### Colors (match AirSaas theme)
```python
MOOD_COLORS = {
    'good': RGBColor(0x03, 0xE2, 0x6B),       # Green
    'issues': RGBColor(0xFF, 0xD4, 0x3B),     # Yellow
    'complicated': RGBColor(0xFF, 0x92, 0x2B), # Orange
    'blocked': RGBColor(0xFF, 0x0A, 0x55),    # Red
}
```

### Mood Labels (French)
```python
MOOD_LABELS = {
    'good': 'Tout va bien',
    'issues': 'Quelques problèmes',
    'complicated': "C'est compliqué",
    'blocked': 'Bloqué',
}
```

---

## Error Handling

- Log errors to `tracking/CLAUDE_ERRORS.md`
- Handle missing data gracefully (empty placeholders)
- Validate data types before formatting
- Catch and report python-pptx exceptions

---

## Checklist Before Generating

- [ ] Template verification passes (`--verify` shows OK)
- [ ] `EXPECTED_SHAPE_POSITIONS` matches current template
- [ ] Using `duplicate_slide()` for project slides
- [ ] Deleting placeholders on Summary/Data Notes
- [ ] Using correct API field names
- [ ] Font sizes appropriate for content density
- [ ] Leading newlines where template has visual titles

## Template Sync Workflow

When template changes:

1. **Run analysis**: `python3 scripts/generate_ppt.py --analyze`
2. **Run `/mapping`**: This updates `EXPECTED_SHAPE_POSITIONS` and field mappings
3. **Verify**: `python3 scripts/generate_ppt.py --verify`
4. **Generate**: `/ppt-skill` or `python3 scripts/generate_ppt.py`
