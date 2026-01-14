# Interactive Field Mapping

Map AirSaas API fields to PPT template placeholders through interactive conversation.

**IMPORTANT:** This command MUST be interactive. DO NOT just display the current mapping status. You MUST follow each phase and ASK the user questions using the AskUserQuestion tool.

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
