#!/usr/bin/env python3
"""
PPT Generator for AirSaas Flash Reports
Uses python-pptx to generate PowerPoint presentations from fetched data.
Uses Systra template when available.
"""

import json
import os
import copy
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

# Paths
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'templates', 'systra_template.pptx')

# Colors from AirSaas moods
MOOD_COLORS = {
    'good': RGBColor(0x03, 0xE2, 0x6B),       # Green
    'issues': RGBColor(0xFF, 0xD4, 0x3B),     # Yellow
    'complicated': RGBColor(0xFF, 0x92, 0x2B), # Orange
    'blocked': RGBColor(0xFF, 0x0A, 0x55),    # Red
}

# Status colors
STATUS_COLORS = {
    'in_progress': RGBColor(0x00, 0x99, 0xFF),  # Blue
    'ongoing_sourcing': RGBColor(0x00, 0x99, 0xFF),
    'finished': RGBColor(0x03, 0xE2, 0x6B),     # Green
    'backlog': RGBColor(0x99, 0x99, 0x99),      # Gray
    'ideation': RGBColor(0xBB, 0x99, 0xFF),     # Purple
    'solution_chosen': RGBColor(0x00, 0xCC, 0xCC), # Teal
}

# French labels for moods
MOOD_LABELS = {
    'good': 'Tout va bien',
    'issues': 'Quelques problèmes',
    'complicated': "C'est compliqué",
    'blocked': 'Bloqué',
}

# Track unfilled fields for Data Notes slide
UNFILLED_FIELDS = []

# Template shape name mappings (from analysis)
# Slide 1 shapes and their purpose:
SLIDE1_SHAPES = {
    'title': 'Google Shape;671;p1',           # "Project review :"
    'mood_comment': 'Google Shape;678;p1',    # Top box "xxx" - mood/status comment
    'info': 'Google Shape;677;p1',            # Deployment area, end users, milestones
    'achievements': 'Google Shape;673;p1',    # Bottom left "xxx" - achievements
    'next_steps': 'Google Shape;679;p1',      # Bottom left lower - next steps
    'budget': 'Google Shape;675;p1',          # Right side - Build BAC/Actual/EAC
    'made': 'Google Shape;676;p1',            # Right side top - Made achievements
    'risks': 'Google Shape;680;p1',           # Right side middle - risks/issues
    'date': 'Google Shape;681;p1',            # Date "dd/mm/yy"
}

# Expected shape positions from systra_template.pptx (for verification)
# Format: (x, y) in inches, rounded to 1 decimal
EXPECTED_SHAPE_POSITIONS = {
    'title': (0.4, 0.2),           # "Project review :"
    'date': (8.9, 0.1),            # "dd/mm/yy"
    'mood_status': (0.4, 0.9),     # Status/Mood/Owner
    'info': (0.4, 2.1),            # Start/End dates, Milestones
    'achievements': (0.4, 3.0),    # Description & Expected Benefits
    'next_steps': (0.4, 4.4),      # Next steps
    'made': (5.1, 0.9),            # Completed milestones
    'risks': (5.1, 2.3),           # Risk level, Attention points
    'budget': (5.1, 4.2),          # Budget BAC/Actual/EAC
}

# Tolerance for position matching (in inches)
POSITION_TOLERANCE = 0.3


def verify_template(template_path):
    """Verify that the template matches expected shape positions.

    Returns:
        tuple: (is_valid, warnings, detected_positions)
        - is_valid: True if template is compatible
        - warnings: List of warning messages
        - detected_positions: Dict of detected shape positions
    """
    warnings = []
    detected_positions = {}

    try:
        prs = Presentation(template_path)
    except Exception as e:
        return False, [f"Cannot open template: {e}"], {}

    if len(prs.slides) < 1:
        return False, ["Template has no slides"], {}

    # Analyze slide 0 (Project Review template)
    slide = prs.slides[0]

    # Build map of all text shapes with their positions
    shape_positions = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            x = round(shape.left.inches, 1)
            y = round(shape.top.inches, 1)
            text_preview = shape.text[:30].replace('\n', ' ') if shape.text else "(empty)"
            shape_positions.append({
                'name': shape.name,
                'x': x,
                'y': y,
                'text': text_preview
            })

    # Check each expected position
    matched_count = 0
    for field_name, (expected_x, expected_y) in EXPECTED_SHAPE_POSITIONS.items():
        # Find closest shape to expected position
        best_match = None
        best_distance = float('inf')

        for sp in shape_positions:
            dist = ((sp['x'] - expected_x) ** 2 + (sp['y'] - expected_y) ** 2) ** 0.5
            if dist < best_distance:
                best_distance = dist
                best_match = sp

        if best_match and best_distance <= POSITION_TOLERANCE:
            detected_positions[field_name] = (best_match['x'], best_match['y'])
            matched_count += 1

            # Warn if position differs slightly
            if best_distance > 0.1:
                warnings.append(
                    f"  {field_name}: Expected ({expected_x}, {expected_y}), "
                    f"found ({best_match['x']}, {best_match['y']}) - offset: {best_distance:.2f}\""
                )
        else:
            detected_positions[field_name] = None
            warnings.append(
                f"  {field_name}: Expected at ({expected_x}, {expected_y}) - NOT FOUND"
            )

    # Determine if valid (at least 70% of shapes found)
    match_ratio = matched_count / len(EXPECTED_SHAPE_POSITIONS)
    is_valid = match_ratio >= 0.7

    # Add summary
    if warnings:
        warnings.insert(0, f"Template verification: {matched_count}/{len(EXPECTED_SHAPE_POSITIONS)} shapes matched")

    if not is_valid:
        warnings.append(
            f"\n⚠️  Template may have changed! Only {match_ratio:.0%} of expected shapes found."
        )
        warnings.append(
            "   Consider running template analysis to update EXPECTED_SHAPE_POSITIONS."
        )

    return is_valid, warnings, detected_positions


def analyze_template(template_path):
    """Analyze template and print all shape positions.

    Use this to update EXPECTED_SHAPE_POSITIONS after template changes.
    """
    prs = Presentation(template_path)

    print(f"\n{'='*60}")
    print(f"Template Analysis: {os.path.basename(template_path)}")
    print(f"{'='*60}")
    print(f"Dimensions: {prs.slide_width.inches:.2f}\" x {prs.slide_height.inches:.2f}\"")
    print(f"Slides: {len(prs.slides)}")
    print(f"Layouts: {len(prs.slide_layouts)}")

    for slide_idx, slide in enumerate(prs.slides):
        print(f"\n--- Slide {slide_idx} ---")
        for shape in slide.shapes:
            if shape.has_text_frame:
                x = round(shape.left.inches, 2)
                y = round(shape.top.inches, 2)
                w = round(shape.width.inches, 2)
                h = round(shape.height.inches, 2)
                text = shape.text[:40].replace('\n', '↵') if shape.text else "(empty)"
                print(f"  ({x}, {y}) [{w}x{h}] {shape.name}")
                print(f"         Text: \"{text}\"")

    print(f"\n{'='*60}")
    print("To update EXPECTED_SHAPE_POSITIONS, copy positions from above.")
    print(f"{'='*60}\n")


def load_data(data_path):
    """Load fetched project data from JSON file."""
    with open(data_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def find_shape_by_name(slide, name):
    """Find a shape by its name."""
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def set_shape_text(shape, text, font_size=None):
    """Set text on a shape, completely replacing existing content.

    Args:
        shape: The shape to modify
        text: The text to set
        font_size: Optional font size in points (e.g., 9 for Pt(9))
    """
    if shape and shape.has_text_frame:
        tf = shape.text_frame
        text_str = str(text) if text else ""

        # Clear ALL existing text by clearing each paragraph
        # We need to keep the first paragraph (can't delete it) but clear others
        while len(tf.paragraphs) > 1:
            # Remove paragraphs by clearing the XML
            p = tf.paragraphs[-1]._p
            tf._txBody.remove(p)

        # Clear the first (and now only) paragraph
        if tf.paragraphs:
            tf.paragraphs[0].clear()

        # Now set the new text - just assign directly
        # This handles newlines as soft returns within the shape
        if tf.paragraphs:
            tf.paragraphs[0].text = text_str
            # Apply font size if specified
            if font_size is not None:
                tf.paragraphs[0].font.size = Pt(font_size)


def format_date(date_str):
    """Format date string to dd/mm/yy."""
    if not date_str:
        return datetime.now().strftime("%d/%m/%y")
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return dt.strftime("%d/%m/%y")
    except:
        return date_str


def duplicate_slide(prs, slide_index):
    """Duplicate a slide including all its content (shapes, images, formatting).

    Uses internal pptx XML manipulation to properly clone slides.
    """
    from copy import deepcopy
    from pptx.parts.slide import SlidePart
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT

    source_slide = prs.slides[slide_index]

    # Get the slide part and make a copy
    slide_layout = source_slide.slide_layout
    sld_id_lst = prs.slides._sldIdLst

    # Add a new slide using the same layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Clear the new slide's shapes (they come from the layout)
    # We'll copy shapes from the source instead
    spTree = new_slide.shapes._spTree

    # Remove all existing shapes except the first element (usually nvGrpSpPr)
    shapes_to_remove = []
    for child in spTree:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp']:
            shapes_to_remove.append(child)

    for shape_el in shapes_to_remove:
        spTree.remove(shape_el)

    # Copy shapes from source slide
    source_spTree = source_slide.shapes._spTree
    for child in source_spTree:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ['sp', 'pic', 'graphicFrame', 'grpSp', 'cxnSp']:
            new_el = deepcopy(child)
            spTree.append(new_el)

    # Copy image relationships
    source_part = source_slide.part
    new_part = new_slide.part

    for rel in source_part.rels.values():
        if rel.reltype == RT.IMAGE:
            # Add the same image relationship to new slide
            new_part.relate_to(rel._target, RT.IMAGE)

    return new_slide


def populate_project_slide(slide, proj_data, reference_data):
    """Populate a project review slide with data.

    This function finds shapes by position (more reliable than name after duplication)
    and updates their text content.
    """
    project = proj_data.get('project', {})
    resolved = proj_data.get('resolved', {})
    milestones = proj_data.get('milestones', [])
    decisions = proj_data.get('decisions', [])
    attention_points = proj_data.get('attention_points', [])

    # Build a map of shapes by approximate position (in inches, rounded)
    # This is more reliable than names which may change after duplication
    shape_map = {}
    for shape in slide.shapes:
        if shape.has_text_frame:
            # Create a position key (rounded to 1 decimal)
            pos_key = (round(shape.left.inches, 1), round(shape.top.inches, 1))
            shape_map[pos_key] = shape

    # Position-based shape mapping (from template analysis):
    # Title: (0.4, 0.2) - "Project review :"
    # Date: (8.9, 0.1) - "dd/mm/yy"
    # Status/Mood: (0.4, 0.9) - Status info
    # Info: (0.4, 2.1) - Deployment/dates
    # Achievements: (0.4, 3.0) - Description
    # Next steps: (0.4, 4.4) - Next steps
    # Made: (5.1, 0.9) - Completed items
    # Risks: (5.1, 2.3) - Risks
    # Budget: (5.1, 4.2) - Budget info

    # Helper to find shape by approximate position
    # Uses smaller tolerance and finds closest match
    def find_shape_by_pos(target_x, target_y, tolerance=0.15):
        best_match = None
        best_distance = float('inf')

        for (x, y), shape in shape_map.items():
            dist = ((x - target_x) ** 2 + (y - target_y) ** 2) ** 0.5
            if dist <= tolerance and dist < best_distance:
                best_match = shape
                best_distance = dist

        return best_match

    # Font sizes for project review slides (smaller for better fit)
    TITLE_FONT = 14
    CONTENT_FONT = 9
    DATE_FONT = 8

    # Title - "Project review : {project_name}"
    title_shape = find_shape_by_pos(0.4, 0.2)
    if title_shape:
        project_name = project.get('name', 'Unknown Project')
        set_shape_text(title_shape, f"Project review : {project_name}", font_size=TITLE_FONT)

    # Date
    date_shape = find_shape_by_pos(8.9, 0.1)
    if date_shape:
        set_shape_text(date_shape, datetime.now().strftime("%d/%m/%y"), font_size=DATE_FONT)

    # Mood/Status comment (top left area)
    mood_shape = find_shape_by_pos(0.4, 0.9)
    if mood_shape:
        mood = resolved.get('mood', project.get('mood', ''))
        status = resolved.get('status', project.get('status', ''))
        owner = project.get('owner', {})
        owner_name = owner.get('name', '') if isinstance(owner, dict) else str(owner)

        comment = f"Status: {status}\nMood: {mood}\nOwner: {owner_name}"
        set_shape_text(mood_shape, comment, font_size=CONTENT_FONT)

    # Info area - Deployment, end users, milestones ref
    info_shape = find_shape_by_pos(0.4, 2.1)
    if info_shape:
        start_date = project.get('start_date', '')
        end_date = project.get('end_date', '')
        ms_count = len(milestones) if milestones else 0
        completed_ms = len([m for m in milestones if m.get('status') == 'done']) if milestones else 0

        info_text = f"Start: {format_date(start_date)}\nEnd: {format_date(end_date)}\nMilestones: {completed_ms}/{ms_count} completed"
        set_shape_text(info_shape, info_text, font_size=CONTENT_FONT)

    # Achievements (from description or goals) - position (0.4, 3.0)
    # Note: The title "Description & Expected Benefits" is in the template background,
    # so we add a blank line at the start to push content below the title
    achievements_shape = find_shape_by_pos(0.4, 3.0)
    if achievements_shape:
        goals = project.get('goals', [])
        if goals:
            goals_text = "\n".join([f"• {g.get('name', g) if isinstance(g, dict) else g}" for g in goals[:4]])
        else:
            desc = project.get('description_text', '')
            if desc:
                goals_text = desc[:180] + ('...' if len(desc) > 180 else '')
            else:
                goals_text = "No achievements documented"
        # Add blank line at start to avoid overlapping with section title in template
        set_shape_text(achievements_shape, "\n" + goals_text, font_size=CONTENT_FONT)

    # Next steps - position (0.4, 4.4)
    # Note: Title "Next Steps" is in template background
    next_shape = find_shape_by_pos(0.4, 4.4)
    if next_shape:
        upcoming = [m for m in milestones if m.get('status') != 'done'][:3] if milestones else []
        if upcoming:
            next_text = "\n".join([f"• {m.get('name', '')}" for m in upcoming])
        elif decisions:
            next_text = "\n".join([f"• {d.get('title', '')}" for d in decisions[:3]])
        else:
            next_text = "See planning for next steps"
        # Add blank line at start to avoid overlapping with section title
        set_shape_text(next_shape, "\n" + next_text, font_size=CONTENT_FONT)

    # Budget - position (5.1, 4.2)
    budget_shape = find_shape_by_pos(5.1, 4.2)
    if budget_shape:
        bac = project.get('budget_capex_initial')
        actual = project.get('budget_capex_used')
        eac = project.get('budget_capex_landing')
        effort_val = project.get('effort')
        effort_used_val = project.get('effort_used')

        budget_lines = ["Build"]
        budget_lines.append(f"BAC: {bac:,.0f} €" if bac is not None else "BAC: N/A")
        budget_lines.append(f"Actual: {actual:,.0f} €" if actual is not None else "Actual: N/A")
        budget_lines.append(f"EAC: {eac:,.0f} €" if eac is not None else "EAC: N/A")
        if effort_val is not None or effort_used_val is not None:
            effort_str = f"Effort: {effort_val or 'N/A'} j"
            if effort_used_val is not None:
                effort_str += f" (used: {effort_used_val} j)"
            budget_lines.append(effort_str)

        set_shape_text(budget_shape, "\n".join(budget_lines), font_size=CONTENT_FONT)

    # Made achievements - position (5.1, 0.9)
    made_shape = find_shape_by_pos(5.1, 0.9)
    if made_shape:
        completed = [m for m in milestones if m.get('status') == 'done'][:3] if milestones else []
        if completed:
            made_text = "Made:\n" + "\n".join([f"• {m.get('name', '')}" for m in completed])
        else:
            made_text = "Made:\nNo completed milestones yet"
        set_shape_text(made_shape, made_text, font_size=CONTENT_FONT)

    # Risks/Issues - position (5.1, 2.3)
    risks_shape = find_shape_by_pos(5.1, 2.3)
    if risks_shape:
        risk_level = resolved.get('risk', project.get('risk', 'Not set'))
        if attention_points:
            ap_text = "\n".join([f"• {ap.get('title', '')}" for ap in attention_points[:3]])
            risks_text = f"Risk Level: {risk_level}\n\nAttention Points:\n{ap_text}"
        else:
            risks_text = f"Risk Level: {risk_level}\n\nNo attention points"
        set_shape_text(risks_shape, risks_text, font_size=CONTENT_FONT)

    # Clear any remaining placeholder shapes with "xxx" text
    # Also clear shapes at positions we didn't fill
    shapes_to_clear = [
        (0.4, 1.9),  # Empty shape 672
        (0.4, 4.3),  # Empty shape 674
    ]
    for target_x, target_y in shapes_to_clear:
        shape = find_shape_by_pos(target_x, target_y)
        if shape:
            set_shape_text(shape, "")


def generate_from_template(data_path, output_path):
    """Generate PPT using the Systra template.

    This version properly preserves all template styling:
    - Images, borders, colors, backgrounds
    - Shape formatting and positioning
    - Fonts and text styles

    Strategy: Use template slide 0 as source, duplicate it for each project.
    The first project uses the original template slide, subsequent projects
    get copies of that slide with all its visual elements.

    Slide structure per spec section 5.1:
    1. Summary slide - List of all projects with mood/status
    2-N. Project slides - Cloned from template slide 1, with data filled
    Last. Data Notes slide - List of unfilled fields
    """
    from copy import deepcopy

    global UNFILLED_FIELDS
    UNFILLED_FIELDS = []  # Reset for new generation

    # TEMPLATE VERIFICATION (per spec: template is source of truth)
    print("Verifying template compatibility...")
    is_valid, warnings, detected_positions = verify_template(TEMPLATE_PATH)

    if warnings:
        for w in warnings:
            print(w)

    if not is_valid:
        print("\n❌ Template verification FAILED!")
        print("   The template structure has changed significantly.")
        print("   Run: python3 scripts/generate_ppt.py --analyze")
        print("   Then update EXPECTED_SHAPE_POSITIONS in the script.")
        print("\n   Proceeding anyway, but output may have misplaced content.\n")
    else:
        print("✓ Template verified OK\n")

    data = load_data(data_path)

    # Load template
    prs = Presentation(TEMPLATE_PATH)

    # Template has 3 slides:
    # 0: Project Review main (this is our template for project slides)
    # 1: Budget table
    # 2: Planning

    reference_data = data.get('reference_data', {})
    projects = data.get('projects', [])

    if not projects:
        print("No projects found in data file")
        return None

    print(f"Template loaded: {len(prs.slides)} slides")
    print(f"Processing {len(projects)} projects...")

    # Delete the Budget and Planning template slides FIRST (index 1 and 2)
    # Do this before any other operations to avoid index confusion
    for slide_idx in [2, 1]:  # Delete in reverse order
        if slide_idx < len(prs.slides):
            rId = prs.slides._sldIdLst[slide_idx].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[slide_idx]

    # Now we have only slide 0 (Project Review template)
    # First, populate it with the first project data
    print(f"  Slide 1: {projects[0].get('project', {}).get('name', 'Unknown')}")
    populate_project_slide(prs.slides[0], projects[0], reference_data)
    track_project_unfilled(projects[0].get('project', {}))

    # For additional projects, duplicate slide 0
    for idx, proj_data in enumerate(projects[1:], start=1):
        project = proj_data.get('project', {})
        project_name = project.get('name', 'Unknown')
        print(f"  Slide {idx + 1}: {project_name}")

        # Duplicate the template slide (slide 0)
        new_slide = duplicate_slide(prs, 0)

        # Fill with project data
        populate_project_slide(new_slide, proj_data, reference_data)
        track_project_unfilled(project)

    # Add Summary slide at the end, then move to beginning
    print("Creating Summary slide...")
    create_summary_slide_on_template(prs, projects, reference_data)

    # Add Data Notes slide at the very end
    print("Creating Data Notes slide...")
    create_data_notes_slide(prs, UNFILLED_FIELDS)

    # Reorder: move Summary from second-to-last to position 0
    move_slide_to_position(prs, len(prs.slides) - 2, 0)

    # Save
    prs.save(output_path)
    print(f"\n✓ PPT saved to: {output_path}")
    print(f"  Total slides: {len(prs.slides)}")
    print(f"  - Summary: 1")
    print(f"  - Project slides: {len(projects)}")
    print(f"  - Data Notes: 1")
    print(f"  Unfilled fields tracked: {len(UNFILLED_FIELDS)}")

    return output_path


def track_project_unfilled(project):
    """Track unfilled fields for a project."""
    project_name = project.get('name', 'Unknown')
    if not project.get('description_text'):
        track_unfilled_field(project_name, "Description", "Non renseigné")
    if not project.get('end_date'):
        track_unfilled_field(project_name, "End date", "Non renseigné")
    if project.get('budget_capex_initial') is None and project.get('budget_capex_used') is None:
        track_unfilled_field(project_name, "Budget", "Non renseigné dans AirSaas")


def move_slide_to_position(prs, from_index, to_index):
    """Move a slide from one position to another."""
    slides = prs.slides._sldIdLst
    slide_to_move = slides[from_index]

    # Remove from current position
    del slides[from_index]

    # Insert at new position
    slides.insert(to_index, slide_to_move)


def delete_placeholder_shapes(slide):
    """Delete all placeholder shapes from a slide.

    This completely removes the shapes rather than just clearing their text,
    which prevents "Double click to edit" from appearing.
    """
    shapes_to_delete = []
    for shape in slide.shapes:
        if shape.is_placeholder:
            shapes_to_delete.append(shape)

    for shape in shapes_to_delete:
        sp = shape._element
        sp.getparent().remove(sp)


def create_summary_slide_on_template(prs, projects, reference_data):
    """Create summary slide using a clean layout.

    Uses layout index 3 (Texte) which has minimal pre-existing shapes.
    """
    # Use "Texte" layout (index 3) - cleaner than Project_Card
    slide_layout = prs.slide_layouts[3] if len(prs.slide_layouts) > 3 else prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    # DELETE all placeholder shapes completely (not just clear text)
    # This prevents "Double click to edit" from appearing
    delete_placeholder_shapes(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(9), Inches(0.4))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Portfolio Flash Report"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Subtitle with date
    subtitle_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.5), Inches(9), Inches(0.25))
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"Projets Vitaux CODIR - {datetime.now().strftime('%d/%m/%Y')}"
    p.font.size = Pt(11)
    p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Projects count
    count_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.8), Inches(9), Inches(0.25))
    tf = count_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"{len(projects)} projets"
    p.font.size = Pt(10)
    p.font.italic = True

    # Table header - adjusted column positions for better alignment
    header_y = 1.1
    col_widths = [0.7, 3.0, 1.3, 1.5, 1.3]
    col_x = [0.3, 1.0, 4.0, 5.4, 7.0]

    headers = ["ID", "Projet", "Status", "Mood", "Owner"]
    for i, header in enumerate(headers):
        box = slide.shapes.add_textbox(Inches(col_x[i]), Inches(header_y), Inches(col_widths[i]), Inches(0.25))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = header
        p.font.size = Pt(8)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Projects list
    row_y = header_y + 0.28
    row_height = 0.28

    for proj_data in projects[:14]:  # Can fit more rows now
        project = proj_data.get('project', {})
        resolved = proj_data.get('resolved', {})

        # Project ID
        short_id = project.get('short_id', '')
        box = slide.shapes.add_textbox(Inches(col_x[0]), Inches(row_y), Inches(col_widths[0]), Inches(row_height))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = short_id
        p.font.size = Pt(7)
        p.font.bold = True

        # Project Name
        name = project.get('name', 'Unknown')[:35]
        if len(project.get('name', '')) > 35:
            name += '...'
        box = slide.shapes.add_textbox(Inches(col_x[1]), Inches(row_y), Inches(col_widths[1]), Inches(row_height))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = name
        p.font.size = Pt(7)

        # Status
        status = resolved.get('status', project.get('status', ''))
        box = slide.shapes.add_textbox(Inches(col_x[2]), Inches(row_y), Inches(col_widths[2]), Inches(row_height))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = status.replace('_', ' ').title() if status else '-'
        p.font.size = Pt(7)
        if status in STATUS_COLORS:
            p.font.color.rgb = STATUS_COLORS[status]

        # Mood with color
        mood = resolved.get('mood', project.get('mood', ''))
        mood_label = MOOD_LABELS.get(mood, mood) if mood else '-'
        box = slide.shapes.add_textbox(Inches(col_x[3]), Inches(row_y), Inches(col_widths[3]), Inches(row_height))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = mood_label
        p.font.size = Pt(7)
        if mood in MOOD_COLORS:
            p.font.color.rgb = MOOD_COLORS[mood]
            p.font.bold = True

        # Owner
        owner = project.get('owner', {})
        owner_name = owner.get('name', '') if isinstance(owner, dict) else ''
        if owner_name:
            parts = owner_name.split()
            owner_short = parts[0] if parts else ''
        else:
            owner_short = '-'
        box = slide.shapes.add_textbox(Inches(col_x[4]), Inches(row_y), Inches(col_widths[4]), Inches(row_height))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = owner_short
        p.font.size = Pt(7)

        row_y += row_height

    if len(projects) > 14:
        box = slide.shapes.add_textbox(Inches(0.3), Inches(row_y + 0.05), Inches(9), Inches(0.2))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = f"... et {len(projects) - 14} autres projets"
        p.font.size = Pt(7)
        p.font.italic = True

    return slide


def update_planning_slide(slide, proj_data):
    """Update planning slide with milestone info."""
    milestones = proj_data.get('milestones', [])

    # Find the planning text shape
    for shape in slide.shapes:
        if shape.has_text_frame:
            if 'Insert planning' in shape.text:
                if milestones:
                    ms_text = "Milestones:\n" + "\n".join([
                        f"{'✓' if m.get('status') == 'done' else '○'} {m.get('name', '')} - {m.get('date', '')}"
                        for m in milestones[:8]
                    ])
                else:
                    ms_text = "No milestones defined for this project"
                set_shape_text(shape, ms_text)
                break


def track_unfilled_field(project_name, field_name, reason):
    """Track unfilled fields for Data Notes slide."""
    global UNFILLED_FIELDS
    UNFILLED_FIELDS.append({
        'project': project_name,
        'field': field_name,
        'reason': reason
    })


def create_data_notes_slide(prs, unfilled_fields):
    """Create Data Notes slide listing unfilled fields.

    According to spec section 5.1:
    Last slide: Data Notes - List of unfilled fields
    """
    # Use "Texte" layout (index 3) - cleaner
    slide_layout = prs.slide_layouts[3] if len(prs.slide_layouts) > 3 else prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    # DELETE all placeholder shapes completely (not just clear text)
    # This prevents "Double click to edit" from appearing
    delete_placeholder_shapes(slide)

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.15), Inches(9), Inches(0.4))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Data Notes"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x00, 0x33, 0x66)

    # Subtitle
    subtitle_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.5), Inches(9), Inches(0.25))
    tf = subtitle_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Champs non remplis / Fields not populated"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Known API limitations (always show these)
    api_limitations = [
        ("Mood comment", "L'API retourne uniquement le code du mood, pas le commentaire"),
        ("Deployment area", "Champ non disponible dans l'API AirSaas"),
        ("End users (actual/target)", "Champ non disponible dans l'API AirSaas"),
    ]

    # Section: API Limitations
    section_y = 0.85
    section_box = slide.shapes.add_textbox(Inches(0.4), Inches(section_y), Inches(9), Inches(0.25))
    tf = section_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Limitations API connues:"
    p.font.size = Pt(9)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x99, 0x33, 0x00)

    row_y = section_y + 0.28
    for field, reason in api_limitations:
        box = slide.shapes.add_textbox(Inches(0.5), Inches(row_y), Inches(9), Inches(0.2))
        tf = box.text_frame
        p = tf.paragraphs[0]
        p.text = f"• {field}: {reason}"
        p.font.size = Pt(7)
        row_y += 0.2

    # Section: Per-project unfilled fields
    if unfilled_fields:
        row_y += 0.15
        section_box = slide.shapes.add_textbox(Inches(0.4), Inches(row_y), Inches(9), Inches(0.25))
        tf = section_box.text_frame
        p = tf.paragraphs[0]
        p.text = "Champs manquants par projet:"
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0x99, 0x33, 0x00)

        row_y += 0.28

        # Group by project
        by_project = {}
        for item in unfilled_fields:
            proj = item['project']
            if proj not in by_project:
                by_project[proj] = []
            by_project[proj].append(f"{item['field']}: {item['reason']}")

        for project_name, fields in list(by_project.items())[:6]:  # Limit
            box = slide.shapes.add_textbox(Inches(0.5), Inches(row_y), Inches(9), Inches(0.18))
            tf = box.text_frame
            p = tf.paragraphs[0]
            p.text = f"{project_name}:"
            p.font.size = Pt(7)
            p.font.bold = True
            row_y += 0.18

            for field_info in fields[:3]:
                box = slide.shapes.add_textbox(Inches(0.7), Inches(row_y), Inches(8.5), Inches(0.16))
                tf = box.text_frame
                p = tf.paragraphs[0]
                p.text = f"- {field_info}"
                p.font.size = Pt(6)
                row_y += 0.16

            row_y += 0.08

    # Footer note
    footer_box = slide.shapes.add_textbox(Inches(0.4), Inches(5.2), Inches(9), Inches(0.2))
    tf = footer_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Source: AirSaas API"
    p.font.size = Pt(6)
    p.font.italic = True
    p.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    return slide


def generate_ppt_simple(data_path, output_path):
    """Generate simple PPT without template (fallback).

    Slide structure per spec section 5.1:
    1. Summary slide - List of all projects with mood/status
    2-N. Project slides - Simple layout for each project
    Last. Data Notes slide - List of unfilled fields
    """
    global UNFILLED_FIELDS
    UNFILLED_FIELDS = []  # Reset for new generation

    data = load_data(data_path)

    # Create presentation (16:9)
    prs = Presentation()
    prs.slide_width = Inches(10)  # Match template
    prs.slide_height = Inches(5.625)

    projects = data.get('projects', [])
    reference_data = data.get('reference_data', {})

    # STEP 1: Create Summary slide first
    print("Creating Summary slide...")
    create_summary_slide(prs, projects, reference_data)

    # STEP 2: Add slides for each project
    print(f"Creating slides for {len(projects)} projects...")

    for proj_data in projects:
        project = proj_data.get('project', {})
        resolved = proj_data.get('resolved', {})
        project_name = project.get('name', 'Unknown')

        # Add a simple slide
        slide_layout = prs.slide_layouts[6]  # Blank
        slide = prs.slides.add_slide(slide_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.17), Inches(8), Inches(0.5))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"Project Review: {project_name}"
        p.font.size = Pt(20)
        p.font.bold = True

        # Content
        content_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.8), Inches(9), Inches(4))
        tf = content_box.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"Status: {resolved.get('status', project.get('status', ''))}"
        p.font.size = Pt(12)

        p = tf.add_paragraph()
        p.text = f"Mood: {resolved.get('mood', project.get('mood', ''))}"
        p.font.size = Pt(12)

        owner = project.get('owner', {})
        owner_name = owner.get('name', '') if isinstance(owner, dict) else ''
        p = tf.add_paragraph()
        p.text = f"Owner: {owner_name}"
        p.font.size = Pt(12)

        p = tf.add_paragraph()
        p.text = f"Timeline: {project.get('start_date', '')} → {project.get('end_date', '')}"
        p.font.size = Pt(12)

        desc = project.get('description_text', '')
        if desc:
            p = tf.add_paragraph()
            p.text = ""
            p = tf.add_paragraph()
            p.text = desc[:400] + ('...' if len(desc) > 400 else '')
            p.font.size = Pt(10)

        # Track unfilled fields
        if not desc:
            track_unfilled_field(project_name, "Description", "Non renseigné")
        if not project.get('end_date'):
            track_unfilled_field(project_name, "End date", "Non renseigné")
        if project.get('budget_capex_initial') is None and project.get('budget_capex_used') is None:
            track_unfilled_field(project_name, "Budget", "Non renseigné dans AirSaas")

    # STEP 3: Create Data Notes slide at the end
    print("Creating Data Notes slide...")
    create_data_notes_slide(prs, UNFILLED_FIELDS)

    prs.save(output_path)
    print(f"✓ PPT saved to: {output_path}")
    print(f"  Total slides: {len(prs.slides)}")
    print(f"  - Summary: 1")
    print(f"  - Project slides: {len(projects)}")
    print(f"  - Data Notes: 1")
    print(f"  Unfilled fields tracked: {len(UNFILLED_FIELDS)}")

    return output_path


def generate_ppt(data_path, output_path):
    """Generate PPT - uses template if available, otherwise creates simple version."""
    if os.path.exists(TEMPLATE_PATH):
        print(f"Using template: {TEMPLATE_PATH}")
        return generate_from_template(data_path, output_path)
    else:
        print("Template not found, generating simple PPT")
        return generate_ppt_simple(data_path, output_path)


if __name__ == "__main__":
    import sys

    # Check for --analyze flag
    if len(sys.argv) > 1 and sys.argv[1] == '--analyze':
        if os.path.exists(TEMPLATE_PATH):
            analyze_template(TEMPLATE_PATH)
        else:
            print(f"Error: Template not found at {TEMPLATE_PATH}")
        sys.exit(0)

    # Check for --verify flag
    if len(sys.argv) > 1 and sys.argv[1] == '--verify':
        if os.path.exists(TEMPLATE_PATH):
            print("Running template verification...\n")
            is_valid, warnings, positions = verify_template(TEMPLATE_PATH)
            for w in warnings:
                print(w)
            if is_valid:
                print("\n✓ Template is compatible")
            else:
                print("\n❌ Template has compatibility issues")
            sys.exit(0 if is_valid else 1)
        else:
            print(f"Error: Template not found at {TEMPLATE_PATH}")
            sys.exit(1)

    # Find latest data file
    data_dir = os.path.join(BASE_DIR, "data")
    data_files = sorted([f for f in os.listdir(data_dir) if f.endswith('_projects.json')], reverse=True)

    if not data_files:
        print("Error: No data files found. Run /fetch first.")
        sys.exit(1)

    data_path = os.path.join(data_dir, data_files[0])
    print(f"Using data: {data_files[0]}")

    # Output path
    today = datetime.now().strftime("%Y-%m-%d")
    output_path = os.path.join(BASE_DIR, "outputs", f"{today}_portfolio_skill.pptx")

    generate_ppt(data_path, output_path)
