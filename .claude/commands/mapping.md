# Interactive Field Mapping

Map AirSaas API fields to PPT template placeholders through interactive conversation.

**IMPORTANT:** This command MUST be interactive. DO NOT just display the current mapping status. You MUST follow each phase and ASK the user questions using the AskUserQuestion tool.

## Key Features

### Visual Layout Awareness
The template has **section titles in the slide layout** (not in content shapes). The mapping system now:
- Tracks layout title positions separately from content areas
- Preserves proper spacing between titles and content
- Uses bullet points (▪) for list items matching template style

### Paragraph Structure Preservation
Content shapes support structured paragraphs:
- **Title paragraphs**: No bullet, bold text (e.g., "Build", "Made :")
- **Content paragraphs**: With bullets, normal text
- Automatic bullet formatting using `set_shape_text_with_structure()`

## Quick Reference - Script Commands

| Command | Purpose |
|---------|---------|
| `python3 scripts/generate_ppt.py --analyze` | Full template analysis with shape positions |
| `python3 scripts/generate_ppt.py --analyze -v` | Verbose analysis (includes non-text shapes) |
| `python3 scripts/generate_ppt.py --verify` | Verify template matches expected positions |
| `python3 scripts/generate_ppt.py --export-mapping` | Export all shapes to JSON file |

---

## Instructions

### Phase 0: COMPREHENSIVE Template Analysis (REQUIRED FIRST)

**CRITICAL:** Before ANY mapping, perform exhaustive template analysis.

1. **Export template mapping to JSON** (creates reference file):
   ```bash
   python3 scripts/generate_ppt.py --export-mapping
   ```
   This creates `config/template_shapes.json` with ALL shape details.

2. **Run full template analysis**:
   ```bash
   python3 scripts/generate_ppt.py --analyze
   ```

   **MUST CAPTURE from output:**
   - Slide dimensions
   - Total number of slides and layouts
   - EVERY text shape with:
     - Exact position (x, y) in inches
     - Size (width x height)
     - Shape name (e.g., `Google Shape;671;p1`)
     - Current text content (or empty)
   - All image shapes (icons, logos)
   - Placeholder shapes that need content

3. **Run verification against current mapping**:
   ```bash
   python3 scripts/generate_ppt.py --verify
   ```

   **Check for:**
   - ✓ Matched shapes - positions are correct
   - ⚠️ Drifted shapes - positions have shifted
   - ❌ Missing shapes - not found at expected position
   - 🆕 New shapes - exist in template but not in mapping

4. **Document ALL fields** - Create a comprehensive list:

   ```
   TEMPLATE FIELD INVENTORY:
   ========================
   Slide 0 (Project Review):
     Position (0.4, 0.2):  TITLE - "Project review : {name}"
     Position (8.9, 0.1):  DATE - "dd/mm/yy"
     Position (0.4, 1.0):  MOOD_STATUS - Status and mood display
     Position (0.35, 1.9): SCOPE_MILESTONES - Dates and milestone count
     Position (0.4, 2.2):  INFO_AREA - (configured empty)
     Position (0.4, 3.2):  ACHIEVEMENTS - Description text
     Position (0.4, 4.3):  TRENDS - Progress percentages
     Position (0.4, 4.5):  NEXT_STEPS - Pending milestones
     Position (5.1, 1.0):  MADE - Completed milestones
     Position (5.1, 2.5):  RISKS - Risk level display
     Position (5.1, 4.4):  BUDGET - BAC/Actual/EAC values
   ```

5. **Present complete template summary to user** showing:
   - Visual layout diagram (ASCII or description)
   - All detected fields with their current purpose
   - Any fields that appear unmapped or have placeholder text ("xxx")

### Phase 1: Compare with Current Mapping (REQUIRED)

1. **Read current mapping.json**:
   ```bash
   cat config/mapping.json
   ```

2. **Identify discrepancies**:
   - Fields in template but NOT in mapping.json
   - Fields in mapping.json but NOT in template (removed/moved)
   - Position differences between mapping.json and actual template

3. **Report to user**:
   ```
   MAPPING SYNC STATUS:
   ====================
   ✓ title: Mapped and position matches
   ✓ date: Mapped and position matches
   ⚠️ scope_milestones: Position drifted (0.35, 1.9) → (0.36, 1.92)
   ❌ new_field_xyz: In template but NOT mapped
   ```

### Phase 2: API Field Inventory (REQUIRED)

1. **Load Available API Fields from ACTUAL data**:
   Read from `data/*_projects.json` to see REAL field names:
   ```python
   import json
   import glob

   # Find latest data file
   files = sorted(glob.glob('data/*_projects.json'), reverse=True)
   with open(files[0]) as f:
       data = json.load(f)

   # Show actual structure
   project = data['projects'][0]
   print("=== PROJECT FIELDS ===")
   print(json.dumps(project.get('project', {}), indent=2))
   print("\n=== RESOLVED VALUES ===")
   print(json.dumps(project.get('resolved', {}), indent=2))
   print("\n=== MILESTONES ===")
   print(f"Count: {len(project.get('milestones', []))}")
   if project.get('milestones'):
       print(json.dumps(project['milestones'][0], indent=2))
   ```

2. **Complete API Fields Reference**:

   **From project object:**
   - `project.name` - Project name
   - `project.short_id` - Short identifier (e.g., "AQM-P13")
   - `project.description_text` - Plain text description
   - `project.status` - Status code (e.g., "in_progress")
   - `project.mood` - Mood code (e.g., "good", "issues")
   - `project.risk` - Risk code (e.g., "low", "medium", "high")
   - `project.owner.name` - Owner's full name
   - `project.program.name` - Program name
   - `project.start_date` - Start date (ISO format)
   - `project.end_date` - End date (ISO format)
   - `project.budget_capex_initial` - BAC (Budget at Completion)
   - `project.budget_capex_used` - Actual spent
   - `project.budget_capex_landing` - EAC (Estimate at Completion)
   - `project.effort` - Planned effort (days)
   - `project.effort_used` - Actual effort used
   - `project.progress` - Progress percentage (0-100)
   - `project.milestone_progress` - Milestone progress (0-100)
   - `project.gain` - Expected gain value
   - `project.gain_text` - Gain description

   **From resolved object (reference data labels):**
   - `resolved.status` - Human-readable status label
   - `resolved.mood` - Human-readable mood label
   - `resolved.risk` - Human-readable risk label

   **From milestones array:**
   - `milestones[].name` - Milestone name
   - `milestones[].status` - Status ('done', 'todo', 'in-progress')
   - `milestones[].date` - Due date
   - `milestones[].objective` - Objective description

   **From other arrays:**
   - `decisions[].title` - Decision title
   - `attention_points[].title` - Attention point title
   - `members[].user.name` - Team member name
   - `members[].role.name` - Member's role
   - `efforts[].team.name` - Team name
   - `efforts[].effort` - Team's planned effort

3. **NOT Available in API** (document these clearly):
   - Mood comment/explanation text
   - Deployment area
   - End users (actual/target)

### Phase 3: Interactive Mapping (REQUIRED - USE AskUserQuestion)

**YOU MUST ASK THE USER** for each field group using the AskUserQuestion tool.

For EACH template field, ask the user what data to display:

1. **Title field**:
   ```
   Question: "For the TITLE field, what format do you want?"
   Options:
   - "Project review : {name}" (Recommended)
   - "Project: {name}"
   - "{name} - {short_id}"
   - "Custom format"
   ```

2. **Status/Mood field**:
   ```
   Question: "For the MOOD_STATUS area, what information to show?"
   Options:
   - "Status + Mood + Owner" (Recommended)
   - "Status + Mood only"
   - "Status only"
   - "Custom combination"
   ```

3. **Scope & Milestones field**:
   ```
   Question: "For SCOPE & MILESTONES, what to display?"
   Options:
   - "Start/End dates + Milestone count" (Recommended)
   - "Start/End dates only"
   - "Milestone count + Progress %"
   - "Leave empty"
   ```

4. **Info Area field**:
   ```
   Question: "For the INFO AREA (deployment/end users), what to show?"
   Options:
   - "Leave empty (data not available in API)" (Recommended)
   - "Show placeholder text"
   - "Use for custom notes"
   ```

5. **Achievements field**:
   ```
   Question: "For ACHIEVEMENTS section, what content?"
   Options:
   - "Project description (description_text)" (Recommended)
   - "Gain description (gain_text)"
   - "Custom text"
   ```

6. **Trends field**:
   ```
   Question: "For TRENDS section, what to display?"
   Options:
   - "Progress % + Milestone Progress %" (Recommended)
   - "Progress % only"
   - "Milestone Progress % only"
   ```

7. **Next Steps field**:
   ```
   Question: "For NEXT STEPS, what to show?"
   Options:
   - "Pending milestones (up to 3)" (Recommended)
   - "Upcoming decisions"
   - "Mix of milestones and decisions"
   ```

8. **Made field**:
   ```
   Question: "For MADE (completed work), what to show?"
   Options:
   - "Completed milestones" (Recommended)
   - "All milestones with status"
   - "Leave empty"
   ```

9. **Risks field**:
   ```
   Question: "For RISKS section, what to display?"
   Options:
   - "Risk level only" (Recommended)
   - "Risk level + Attention points"
   - "Attention points only"
   ```

10. **Budget field**:
    ```
    Question: "For BUDGET section, what values?"
    Options:
    - "BAC + Actual + EAC (CAPEX)" (Recommended)
    - "BAC + Actual + EAC + Effort"
    - "Only if values exist"
    ```

11. **Date format**:
    ```
    Question: "What date format to use throughout?"
    Options:
    - "dd/mm/yy" (Recommended)
    - "dd/mm/yyyy"
    - "ISO (2024-01-15)"
    ```

### Phase 4: Save Results (REQUIRED)

1. **Update config/mapping.json** with ALL user choices:
   - Update `user_preferences` section
   - Update each field in `slides.project_card`
   - Include shape name, position, source, format, and notes
   - **Add `content_structure` for fields with bullets or inline titles:**

   ```json
   "content_structure": {
     "layout_title_above": "SECTION_TITLE",       // Title from layout (if any)
     "layout_title_position": "(x, y)",           // Position of layout title
     "content_offset_y": 0.10,                    // Vertical space below title
     "first_para_is_title": true,                 // If shape has inline title
     "title_text": "Build",                       // Inline title text
     "use_bullets": true,                         // Enable bullet formatting
     "bullet_char": "▪",                          // Bullet character to use
     "items_as_paragraphs": true,                 // Each item as separate paragraph
     "max_items": 4                               // Maximum items to display
   }
   ```

2. **Update scripts/generate_ppt.py** if positions changed:
   - Update `EXPECTED_SHAPE_POSITIONS` constant
   - Update position arguments in `populate_project_slide()` function
   - Verify field names match API data structure

3. **Update tracking/MISSING_FIELDS.md** with any missing fields noted

4. **Show summary to user**:
   ```
   MAPPING COMPLETE
   ================
   Fields configured: 11
   Positions updated: 2
   New fields added: 0

   Files updated:
   - config/mapping.json
   - scripts/generate_ppt.py (EXPECTED_SHAPE_POSITIONS)
   ```

### Phase 5: Final Verification (CRITICAL)

1. **Run verification**:
   ```bash
   python3 scripts/generate_ppt.py --verify
   ```
   Must show "✓ Template is compatible"

2. **Cross-check with actual data**:
   - Confirm field names in script match API data exactly
   - Test at least one value lookup mentally

3. **Generate test PPT** (optional but recommended):
   ```bash
   python3 scripts/generate_ppt.py
   ```

4. **Report to user**:
   ```
   ✓ Verification complete
   ✓ Template positions synchronized
   ✓ Field names validated against API data

   Ready to generate PPT with /ppt-skill
   ```

---

## CRITICAL RULES

1. **NEVER skip template analysis** - Run --analyze FIRST
2. **CAPTURE ALL SHAPES** - Don't miss any text shape
3. **VERIFY POSITIONS** - Run --verify after any changes
4. **USE AskUserQuestion** - Ask for EVERY field, never assume
5. **READ ACTUAL DATA** - Get field names from data/*.json, not memory
6. **UPDATE ALL FILES** - mapping.json AND generate_ppt.py must match
7. **DOCUMENT EVERYTHING** - Include notes for each mapping decision
8. **LOG ERRORS** - Update tracking/CLAUDE_ERRORS.md if issues found

---

## PPT Generation Rules (for generate_ppt.py updates)

### Template Preservation (MANDATORY)

1. **Duplicate slides via XML** to preserve:
   - Images (icons, logos, backgrounds)
   - Borders and fills
   - Colors and gradients
   - Font styles

2. **Find shapes by position** (not by name):
   - Shape names change after duplication
   - Use Euclidean distance with tolerance (~0.15 inches)

3. **Handle visual section titles**:
   - If template has titles in background
   - Add leading newline to push content below

### Font Size Guidelines

| Element | Recommended Size |
|---------|------------------|
| Slide title | 14-20pt |
| Section headers | 10-12pt |
| Body content | 8-10pt |
| Table content | 7-8pt |
| Date/metadata | 8pt |
| Dense notes | 6-7pt |

### Content Setting Functions

Use the appropriate function based on content type:

| Function | Use Case | Example |
|----------|----------|---------|
| `set_shape_text()` | Simple text without bullets | Title, date, single-line values |
| `set_shape_text_with_structure()` | Bullet lists, multi-paragraph | SCOPE, NEXT STEPS, BUDGET, MADE |

```python
# Simple text (no bullets)
set_shape_text(shape, "Progress: 66%", font_size=9)

# Structured content with bullets
set_shape_text_with_structure(
    shape,
    items=["Item 1", "Item 2", "Item 3"],  # List of strings
    title_text="Build",                      # Optional inline title (no bullet)
    font_size=9,
    use_bullets=True
)
```

### Layout vs Shape Titles

The template has two types of section titles:

1. **Layout titles** (in slide layout, not editable):
   - "SCOPE & MILESTONES", "NEXT STEPS", "BUDGET", etc.
   - Appear above content shapes
   - DO NOT write over these - content must appear below them

2. **Shape inline titles** (in content shape, first paragraph):
   - "Build", "Made :"
   - Set via `title_text` parameter
   - Automatically formatted without bullet, bold

### Shape Overlap Handling

**CRITICAL:** Some content shapes physically overlap in the template!

| Overlapping Shapes | Solution |
|--------------------|----------|
| 672 + 677 | Use only 672, clear 677 |
| 674 + 679 | Use only 679, clear 674 |

**Shape 672 (SCOPE & MILESTONES):**
- Layout title overlaps top 0.16" of shape
- Use `set_shape_text_with_title_padding()` with `padding_lines=1`
- First paragraph must be EMPTY

```python
# Shape with layout title overlap
set_shape_text_with_title_padding(
    scope_shape,
    items=["Milestones: 0/5", "Progress: 66%"],
    padding_lines=1,  # Empty first line
    font_size=9,
    use_bullets=True
)

# Overlapping shape - CLEAR IT
set_shape_text(info_shape, '', font_size=9)
```

### Template Color Palette

Colors are defined in `mapping.json._color_palette`:

| Element | Color | Hex |
|---------|-------|-----|
| Section titles | Dark teal | #003D4B |
| Content text | Black | #000000 |
| Date background | Gray-teal | #78989F |
| Mood: good | Green | #03E26B |
| Mood: issues | Yellow | #FFD43B |
| Mood: complicated | Orange | #FF922B |
| Mood: blocked | Red | #FF0A55 |
| Risk: low | Green | #03E26B |
| Risk: medium | Yellow | #FFD43B |
| Risk: high | Red | #FF0A55 |

---

## Style Extraction Rules (CRITICAL)

**IMPORTANT:** The script extracts visual styles from the template at runtime. These rules are defined in `mapping.json._style_extraction_rules` and MUST be updated when the template changes.

### Why Style Extraction?

The script must:
1. **Extract** font names, sizes, colors, and bullet styles from template shapes
2. **Apply uniformly** to ALL generated content of the same type
3. **Never hardcode** visual properties - always read from template

This ensures visual consistency even when templates change.

### Style Types Configuration

Located in `mapping.json._style_extraction_rules.style_types`:

| Style Type | Extract From | Properties | Apply To |
|------------|--------------|------------|----------|
| `title` | Shape at (x, y) | font_name, size, color, bold | Slide titles |
| `date` | Shape at (x, y) | font_name, size, color | Date fields |
| `status` | Shape at (x, y) | font_name, size, color | Status/mood text |
| `inline_title` | Multiple positions | font_name, size, color, bold | "Build", "Made :" |
| `content` | Y positions list | font_name, size, color | Bullet items |
| `bullet` | First colored bullet | char, color | ALL bullets |

### Configuration Structure

```json
"_style_extraction_rules": {
  "style_types": {
    "title": {
      "extract_from_position": {"x": 0.39, "y": 0.17},
      "shape_id": "671",
      "extract": ["font_name", "font_size", "font_color", "bold"]
    },
    "bullet": {
      "skip_colors": ["000000", "FFFFFF"],
      "note": "Takes first non-black/white bullet color"
    },
    "content": {
      "extract_from_y_positions": [1.93, 3.16, 4.53, 2.48],
      "shape_ids": ["672", "673", "679", "680"]
    }
  },
  "defaults": {
    "font_name": "Lato",
    "font_size_pt": 9,
    "bullet_color": "D32427"
  }
}
```

### Application Rules

From `mapping.json._style_extraction_rules.application_rules`:

1. **NEVER hardcode** font names, sizes, or colors in script
2. **ALWAYS extract** from template at runtime
3. **Bullet color MUST be uniform** across all slides
4. **Font family MUST match** template exactly
5. When template changes, re-extract styles automatically

### When Running /mapping

**Phase 0 MUST also:**

1. **Extract bullet color** from template:
   - Find first shape with non-black/white bullet
   - Record the color (e.g., `#D32427` for red)

2. **Extract font styles** for each style type:
   - Title font (from shape 671)
   - Content font (from bullet shapes)
   - Date font (from shape 681)

3. **Update `_style_extraction_rules`** in mapping.json:
   - Update positions if shapes moved
   - Update defaults if colors changed
   - Update `skip_colors` if needed

### Example: Updating for New Template

If template changes to use blue bullets (#0066CC):

```json
"_style_extraction_rules": {
  "style_types": {
    "bullet": {
      "skip_colors": ["000000", "FFFFFF"],
      "note": "Blue bullets in new template"
    }
  },
  "defaults": {
    "bullet_color": "0066CC"
  }
}
```

---

## Example Session

```
Claude: Running comprehensive template analysis...

        $ python3 scripts/generate_ppt.py --analyze

        COMPREHENSIVE TEMPLATE ANALYSIS
        ================================
        File: ProjectCardAndFollowUp.pptx
        Dimensions: 10.00" x 5.62"
        Total Slides: 3

        SLIDE 0 (Project Review):
        📝 TEXT SHAPES (11):
          ▸ Google Shape;671;p1
            Position: (0.4, 0.2) | Size: 5.0" x 0.3"
            Text: "Project review :"

          ▸ Google Shape;681;p1
            Position: (8.9, 0.1) | Size: 1.0" x 0.2"
            Text: "dd/mm/yy"
          ... (all other shapes)

        Now let me verify against current mapping...

        $ python3 scripts/generate_ppt.py --verify

        VERIFICATION REPORT
        ===================
        ✓ 11/11 shapes matched
        Template is compatible

Claude: Template analysis complete. I found 11 mappable fields.

        [Uses AskUserQuestion]
        "For the TITLE field, what format?"

User: [Selects "Project review : {name}"]

Claude: Perfect. Now for the MOOD_STATUS area...
        [Uses AskUserQuestion]

[... continues for each field ...]

Claude: Mapping complete! Updating files...

        ✓ config/mapping.json updated
        ✓ scripts/generate_ppt.py positions verified
        ✓ Final verification passed

        Ready to generate PPT with /ppt-skill
```
