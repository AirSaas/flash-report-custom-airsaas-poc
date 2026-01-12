# Generate PPT via python-pptx

Generate a PowerPoint presentation using python-pptx for programmatic control.

## Instructions

### Phase 1: Load Data

1. **Find Latest Fetched Data**
   - Look in `data/` folder for most recent `{date}_projects.json`
   - If no data found, prompt user to run `/fetch` first

2. **Load Mapping Configuration**
   - Read `config/mapping.json`
   - Identify fields and their sources/transforms

3. **Load Template (if available)**
   - Check `templates/` for reference PPTX
   - Use as basis for slide layouts and styling

4. **Load Reference Data**
   - Extract moods, statuses, risks from fetched data
   - Build lookup tables for label resolution

### Phase 2: Setup python-pptx

1. **Verify Installation**
   ```bash
   pip install python-pptx
   ```

2. **Create Python Script**
   Create a script in `scripts/` or execute directly:

   ```python
   from pptx import Presentation
   from pptx.util import Inches, Pt
   from pptx.dml.color import RgbColor
   from pptx.enum.text import PP_ALIGN
   import json
   from datetime import datetime
   ```

### Phase 3: Generate Presentation

1. **Create Presentation Object**
   ```python
   # Use template if available, otherwise blank
   if template_path:
       prs = Presentation(template_path)
   else:
       prs = Presentation()
       prs.slide_width = Inches(13.333)  # 16:9
       prs.slide_height = Inches(7.5)
   ```

2. **Add Summary Slide**
   ```python
   slide_layout = prs.slide_layouts[5]  # Blank or Title
   slide = prs.slides.add_slide(slide_layout)

   # Add title
   title = slide.shapes.title
   title.text = "Portfolio Flash Report"

   # Add date
   # Add project table with status/mood
   ```

3. **Add Per-Project Slides**
   For each project, following mapping.json structure:

   **Project Card Slide:**
   ```python
   slide = prs.slides.add_slide(slide_layout)
   # Add project name as title
   # Add status, mood, risk indicators
   # Add owner, program info
   # Add budget, timeline
   # Add achievements/goals
   ```

   **Progress Slide:**
   ```python
   slide = prs.slides.add_slide(slide_layout)
   # Add completion percentage
   # Add milestones timeline
   # Add KPIs
   ```

   **Planning Slide:**
   ```python
   slide = prs.slides.add_slide(slide_layout)
   # Add team efforts table
   # Add decisions list
   # Add attention points
   ```

4. **Add Data Notes Slide**
   ```python
   slide = prs.slides.add_slide(slide_layout)
   # Title: "Data Notes"
   # List unfilled fields
   ```

5. **Apply Transforms**
   Implement transform functions:
   ```python
   def resolve_status_label(code, statuses):
       return next((s['label'] for s in statuses if s['code'] == code), code)

   def format_date(iso_date):
       if not iso_date:
           return ""
       return datetime.fromisoformat(iso_date).strftime("%d/%m/%Y")

   def calculate_completion(milestones):
       if not milestones:
           return ""
       completed = sum(1 for m in milestones if m.get('is_completed'))
       return f"{int(completed / len(milestones) * 100)}%"
   ```

### Phase 4: Save and Report

1. **Save Presentation**
   ```python
   output_path = f"outputs/{datetime.now().strftime('%Y-%m-%d')}_portfolio_skill.pptx"
   prs.save(output_path)
   ```

2. **Report Results**
   ```
   PPT Generated Successfully!

   Output: outputs/{date}_portfolio_skill.pptx
   Projects: {count}
   Slides: {total_slides}

   Unfilled Fields:
   - {field1} in {project1}
   - {field2} in {project2}
   ```

## Styling Guidelines

### Colors (match AirSaas theme)
```python
COLORS = {
    'primary': RgbColor(0x00, 0x72, 0xC6),    # Blue
    'success': RgbColor(0x28, 0xA7, 0x45),    # Green
    'warning': RgbColor(0xFF, 0xC1, 0x07),    # Yellow
    'danger': RgbColor(0xDC, 0x35, 0x45),     # Red
    'text': RgbColor(0x33, 0x33, 0x33),       # Dark gray
}
```

### Fonts
```python
FONTS = {
    'title': ('Calibri', Pt(28)),
    'heading': ('Calibri', Pt(20)),
    'body': ('Calibri', Pt(14)),
    'small': ('Calibri', Pt(11)),
}
```

### Layout
- Margins: 0.5 inches
- Title at top
- Content area below
- Tables: full width with headers

## Error Handling

- Log errors to `tracking/CLAUDE_ERRORS.md`
- Handle missing data gracefully (empty placeholders)
- Validate data types before formatting
- Catch and report python-pptx exceptions
