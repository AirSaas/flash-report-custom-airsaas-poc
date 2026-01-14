# Interactive Field Mapping

Map AirSaas API fields to PPT template placeholders through interactive conversation.

**IMPORTANT:** This command MUST be interactive. DO NOT just display the current mapping status. You MUST follow each phase and ASK the user questions using the AskUserQuestion tool.

## Instructions

### Phase 0: Template Verification (REQUIRED FIRST)

**CRITICAL:** Before any mapping, verify and sync template positions with `generate_ppt.py`.

1. **Run template verification**:
   ```bash
   python3 scripts/generate_ppt.py --verify
   ```

2. **If verification fails or template changed**, run analysis:
   ```bash
   python3 scripts/generate_ppt.py --analyze
   ```

3. **Update `EXPECTED_SHAPE_POSITIONS` in `generate_ppt.py`** if positions changed:
   - Read the analysis output
   - Update the constant with new positions
   - This ensures `/ppt-skill` will place content correctly

This step syncs the script with any template changes before mapping fields.

### Phase 1: Template Analysis (REQUIRED)

1. **Check for Template File**
   - Look in `templates/` folder for PPT files
   - If no template exists, ask user to provide one or describe desired structure

2. **Analyze Template Structure with python-pptx**
   Run Python code to extract ALL shapes from the template:
   ```python
   from pptx import Presentation
   prs = Presentation('templates/YOUR_TEMPLATE.pptx')

   print(f"Slide dimensions: {prs.slide_width.inches} x {prs.slide_height.inches} inches")
   print(f"Number of slides: {len(prs.slides)}")
   print(f"Number of layouts: {len(prs.slide_layouts)}")
   print()

   for slide_idx, slide in enumerate(prs.slides):
       print(f"Slide {slide_idx}:")
       for shape in slide.shapes:
           x = round(shape.left.inches, 2)
           y = round(shape.top.inches, 2)
           w = round(shape.width.inches, 2)
           h = round(shape.height.inches, 2)
           shape_type = type(shape).__name__
           print(f"  - {shape.name} ({shape_type})")
           print(f"    Position: ({x}, {y}) Size: {w}x{h}")
           if shape.has_text_frame:
               text = shape.text[:50].replace('\n', ' | ')
               print(f"    Text: '{text}...'")
   ```

3. **Identify Visual Elements**
   - Count images per slide
   - Note shapes with borders/fills
   - Identify section titles in background (visual elements, not text shapes)

4. **Present Template Summary to User**
   Show the user what placeholders were found in each slide.

### Phase 2: API Field Inventory (REQUIRED)

1. **Load Available API Fields from ACTUAL data**
   Read from `data/*_projects.json` to see REAL field names:
   ```python
   import json
   import glob

   # Find latest data file
   files = sorted(glob.glob('data/*_projects.json'), reverse=True)
   with open(files[0]) as f:
       data = json.load(f)

   # Show actual structure
   print(json.dumps(data['projects'][0], indent=2))
   ```

   **Correct API Fields (from actual data):**
   - `project.name`, `project.short_id`, `project.description_text`
   - `project.status`, `project.mood`, `project.risk` (codes)
   - `project.owner.name`, `project.program.name`
   - `project.start_date`, `project.end_date`
   - `project.budget_capex_initial` (BAC)
   - `project.budget_capex_used` (Actual)
   - `project.budget_capex_landing` (EAC)
   - `project.effort`, `project.effort_used`
   - `project.progress`, `project.milestone_progress`
   - `project.gain_text`
   - `milestones[]`: name, status ('done'/'todo'/'in-progress'), objective
   - `decisions[]`: title, status
   - `attention_points[]`: title, severity, status
   - `members[]`: user.name, role.name
   - `efforts[]`: team.name, effort, effort_used

2. **Present API Fields to User**
   Show organized list with example values from real data

### Phase 3: Interactive Mapping (REQUIRED - USE AskUserQuestion)

**YOU MUST ASK THE USER** for each field group using the AskUserQuestion tool.

For each PPT placeholder group:

1. **Use AskUserQuestion tool** to propose matches:
   ```
   Question: "Para el campo 'Budget' del template, ¿qué campos quieres usar?"
   Options:
   - "BAC: budget_capex_initial, Actual: budget_capex_used, EAC: budget_capex_landing"
   - "Solo mostrar effort (días)"
   - "Dejarlo vacío"
   - "Otra configuración"
   ```

2. **Wait for user response before continuing**

3. **Confirm transforms needed:**
   ```
   Question: "¿Qué formato para las fechas?"
   Options:
   - "dd/mm/yyyy"
   - "dd/mm/yy"
   - "ISO (2024-01-15)"
   ```

### Phase 4: Save Results (REQUIRED)

1. **Update mapping.json** with the user's choices
2. **Update generate_ppt.py** if field names changed
3. **Update MISSING_FIELDS.md** with any missing fields
4. **Show summary** of what was configured

### Phase 5: Sync Template Positions (REQUIRED)

**CRITICAL: After mapping, ensure `generate_ppt.py` has correct shape positions.**

1. **Analyze current template**:
   ```bash
   python3 scripts/generate_ppt.py --analyze
   ```

2. **Compare with `EXPECTED_SHAPE_POSITIONS`** in `generate_ppt.py`:
   - Read the current constant values
   - Compare with analysis output
   - If different, update the constant

3. **Update `EXPECTED_SHAPE_POSITIONS`** if needed:
   ```python
   # In scripts/generate_ppt.py, update this constant:
   EXPECTED_SHAPE_POSITIONS = {
       'title': (x, y),           # From analysis
       'date': (x, y),
       'mood_status': (x, y),
       # ... etc
   }
   ```

4. **Also update shape positions in `populate_project_slide()`** if they changed:
   - The function uses hardcoded positions like `find_shape_by_pos(0.4, 0.2)`
   - These must match the actual template positions

### Phase 6: Final Verification (REQUIRED)

**CRITICAL: Before finishing, ALWAYS verify the implementation matches the mapping.**

1. **Run verification**:
   ```bash
   python3 scripts/generate_ppt.py --verify
   ```
   Must show "✓ Template is compatible"

2. **Verify field names match mapping.json**:
   ```python
   # Verify these match the actual API field names:
   # - budget_capex_initial (not budget_capex)
   # - budget_capex_used (not budget_actual)
   # - budget_capex_landing (not budget_eac)
   # - status == 'done' for milestones (not is_completed)
   ```

3. **Cross-check with actual data**:
   - Read latest `data/*_projects.json`
   - Confirm field names in script match field names in data
   - Test at least one value lookup mentally

4. **Update CLAUDE_ERRORS.md** if any issues were found

5. **Report to user**: "Verificación completada. Template sincronizado y campos correctos."

## CRITICAL RULES

1. **NEVER skip the interactive phase** - Always ask the user
2. **NEVER assume defaults** without asking
3. **USE AskUserQuestion tool** for each decision
4. **READ ACTUAL DATA** to get correct field names
5. **UPDATE THE SCRIPT** if mapping changes (generate_ppt.py must match mapping.json)
6. **SYNC TEMPLATE POSITIONS** - Run Phase 5 to update EXPECTED_SHAPE_POSITIONS
7. **ALWAYS VERIFY** - Run Phase 6 with `--verify` flag before finishing
8. **LOG ERRORS** - Update tracking/CLAUDE_ERRORS.md if any issues found

## Example Flow

```
Claude: Analizando template en templates/...
        Encontré X slides con estos campos:
        - Slide 0: [shapes found with positions]
        - Slide 1: ...

        [Uses AskUserQuestion]
        "Para el campo 'Budget', ¿cuáles valores quieres mostrar?"

User: [Selects option]

Claude: Perfecto. Ahora para 'Milestones'...
        [Uses AskUserQuestion]

[... continues for each field ...]

Claude: Mapping completado. Guardando en config/mapping.json...
        También actualicé scripts/generate_ppt.py con los campos correctos.
```

---

## PPT Generation Rules (for generate_ppt.py updates)

When updating the script, ensure these rules are followed:

### Template Preservation (MANDATORY)

1. **Duplicate slides via XML** to preserve:
   - Images (icons, logos, backgrounds)
   - Borders and fills
   - Colors and gradients
   - Font styles
   - Shape formatting

2. **Find shapes by position** (not by name):
   - Shape names change after duplication
   - Use Euclidean distance with tolerance (~0.15 inches)
   - Build position map from template analysis

3. **Delete placeholders entirely** on new slides (Summary/Data Notes):
   - Clearing text shows "Double click to edit"
   - Must remove shape element from XML

4. **Handle visual section titles**:
   - If template has titles in background (not in text shapes)
   - Add leading newline to push content below visual title

### Font Size Guidelines

| Element | Recommended Size |
|---------|------------------|
| Slide title | 14-20pt |
| Section headers | 10-12pt |
| Body content | 8-10pt |
| Table content | 7-8pt |
| Date/metadata | 8pt |
| Dense notes | 6-7pt |

### Code Patterns

```python
# Duplicate slide preserving all visuals
def duplicate_slide(prs, slide_index):
    from copy import deepcopy
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    # ... (full implementation in ppt-skill.md)

# Find shape by position
def find_shape_by_pos(slide, target_x, target_y, tolerance=0.15):
    # ... (full implementation in ppt-skill.md)

# Delete placeholders
def delete_placeholder_shapes(slide):
    shapes_to_delete = [s for s in slide.shapes if s.is_placeholder]
    for shape in shapes_to_delete:
        sp = shape._element
        sp.getparent().remove(sp)

# Set text with optional font size
def set_shape_text(shape, text, font_size=None):
    if shape and shape.has_text_frame:
        tf = shape.text_frame
        # Clear existing paragraphs
        while len(tf.paragraphs) > 1:
            tf._txBody.remove(tf.paragraphs[-1]._p)
        if tf.paragraphs:
            tf.paragraphs[0].clear()
            tf.paragraphs[0].text = str(text) if text else ""
            if font_size:
                tf.paragraphs[0].font.size = Pt(font_size)
```
