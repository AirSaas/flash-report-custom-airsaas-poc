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


def set_shape_text(shape, text):
    """Set text on a shape, preserving formatting."""
    if shape and shape.has_text_frame:
        tf = shape.text_frame
        # Clear existing paragraphs but keep first one
        for para in tf.paragraphs:
            para.clear()
        # Set new text
        if tf.paragraphs:
            tf.paragraphs[0].text = str(text) if text else ""


def format_date(date_str):
    """Format date string to dd/mm/yy."""
    if not date_str:
        return datetime.now().strftime("%d/%m/%y")
    try:
        dt = datetime.fromisoformat(date_str.replace('Z', '+00:00'))
        return dt.strftime("%d/%m/%y")
    except:
        return date_str


def clone_slide(prs, template_slide):
    """Clone a slide from template."""
    # Get the layout from the template slide
    slide_layout = template_slide.slide_layout
    new_slide = prs.slides.add_slide(slide_layout)

    # Copy shapes from template
    for shape in template_slide.shapes:
        # Skip placeholders that are already on the layout
        if shape.is_placeholder:
            continue
        # Clone the shape by adding it with same properties
        # This is a simplified approach - complex shapes may need special handling

    return new_slide


def populate_project_slide(slide, proj_data, reference_data):
    """Populate a project review slide with data."""
    project = proj_data.get('project', {})
    resolved = proj_data.get('resolved', {})
    milestones = proj_data.get('milestones', [])
    decisions = proj_data.get('decisions', [])
    attention_points = proj_data.get('attention_points', [])

    # Title - "Project review : {project_name}"
    title_shape = find_shape_by_name(slide, SLIDE1_SHAPES['title'])
    if title_shape:
        project_name = project.get('name', 'Unknown Project')
        set_shape_text(title_shape, f"Project review : {project_name}")

    # Date
    date_shape = find_shape_by_name(slide, SLIDE1_SHAPES['date'])
    if date_shape:
        set_shape_text(date_shape, datetime.now().strftime("%d/%m/%y"))

    # Mood/Status comment (top area)
    mood_shape = find_shape_by_name(slide, SLIDE1_SHAPES['mood_comment'])
    if mood_shape:
        mood = resolved.get('mood', project.get('mood', ''))
        status = resolved.get('status', project.get('status', ''))
        owner = project.get('owner', {})
        owner_name = owner.get('name', '') if isinstance(owner, dict) else str(owner)

        comment = f"Status: {status}\nMood: {mood}\nOwner: {owner_name}"
        set_shape_text(mood_shape, comment)

    # Info area - Deployment, end users, milestones ref
    info_shape = find_shape_by_name(slide, SLIDE1_SHAPES['info'])
    if info_shape:
        start_date = project.get('start_date', '')
        end_date = project.get('end_date', '')
        ms_count = len(milestones) if milestones else 0
        completed_ms = len([m for m in milestones if m.get('is_completed')]) if milestones else 0

        info_text = f"Start: {format_date(start_date)}\nEnd: {format_date(end_date)}\nMilestones: {completed_ms}/{ms_count} completed"
        set_shape_text(info_shape, info_text)

    # Achievements (from description or goals)
    achievements_shape = find_shape_by_name(slide, SLIDE1_SHAPES['achievements'])
    if achievements_shape:
        goals = project.get('goals', [])
        if goals:
            goals_text = "\n".join([f"• {g.get('name', g) if isinstance(g, dict) else g}" for g in goals[:4]])
        else:
            # Use first part of description if no goals
            desc = project.get('description_text', '')
            if desc:
                # Extract objectives from description
                goals_text = desc[:200] + ('...' if len(desc) > 200 else '')
            else:
                goals_text = "No achievements documented"
        set_shape_text(achievements_shape, goals_text)

    # Next steps
    next_shape = find_shape_by_name(slide, SLIDE1_SHAPES['next_steps'])
    if next_shape:
        # Use upcoming milestones as next steps
        upcoming = [m for m in milestones if not m.get('is_completed')][:3] if milestones else []
        if upcoming:
            next_text = "\n".join([f"• {m.get('name', '')}" for m in upcoming])
        elif decisions:
            next_text = "\n".join([f"• {d.get('title', '')}" for d in decisions[:3]])
        else:
            next_text = "See planning for next steps"
        set_shape_text(next_shape, next_text)

    # Budget (not available from API)
    budget_shape = find_shape_by_name(slide, SLIDE1_SHAPES['budget'])
    if budget_shape:
        budget_capex = project.get('budget_capex')
        budget_opex = project.get('budget_opex')
        if budget_capex or budget_opex:
            budget_text = f"Build\nCAPEX: {budget_capex or 'N/A'} K€\nOPEX: {budget_opex or 'N/A'} K€"
        else:
            budget_text = "Build\nBAC: N/A (API unavailable)\nActual: N/A\nEAC: N/A"
        set_shape_text(budget_shape, budget_text)

    # Made achievements
    made_shape = find_shape_by_name(slide, SLIDE1_SHAPES['made'])
    if made_shape:
        completed = [m for m in milestones if m.get('is_completed')][:3] if milestones else []
        if completed:
            made_text = "Made:\n" + "\n".join([f"• {m.get('name', '')}" for m in completed])
        else:
            made_text = "Made:\nNo completed milestones yet"
        set_shape_text(made_shape, made_text)

    # Risks/Issues
    risks_shape = find_shape_by_name(slide, SLIDE1_SHAPES['risks'])
    if risks_shape:
        risk_level = resolved.get('risk', project.get('risk', 'Not set'))
        if attention_points:
            ap_text = "\n".join([f"• {ap.get('title', '')}" for ap in attention_points[:3]])
            risks_text = f"Risk Level: {risk_level}\n\nAttention Points:\n{ap_text}"
        else:
            risks_text = f"Risk Level: {risk_level}\n\nNo attention points"
        set_shape_text(risks_shape, risks_text)


def generate_from_template(data_path, output_path):
    """Generate PPT using the Systra template."""
    data = load_data(data_path)

    # Load template
    prs = Presentation(TEMPLATE_PATH)

    # Template has 3 slides:
    # 0: Project Review main
    # 1: Budget table
    # 2: Planning

    # We'll use slide 0 as the template for each project
    # First, get reference data
    reference_data = data.get('reference_data', {})
    projects = data.get('projects', [])

    if not projects:
        print("No projects found in data file")
        return None

    # Get the template slides
    template_slide = prs.slides[0]

    # Populate first project into existing slide 0
    if projects:
        populate_project_slide(template_slide, projects[0], reference_data)

    # For additional projects, duplicate slide 0 layout and populate
    for proj_data in projects[1:]:
        # Add new slide with same layout
        new_slide = prs.slides.add_slide(template_slide.slide_layout)

        # Copy non-placeholder shapes from template
        project = proj_data.get('project', {})
        resolved = proj_data.get('resolved', {})

        # Since we can't easily clone shapes, we'll add text boxes manually
        # for the additional projects
        add_project_review_content(new_slide, proj_data, reference_data)

    # Update slide 1 (budget) title if it exists
    if len(prs.slides) > 1:
        budget_slide = prs.slides[1]
        title_shape = None
        for shape in budget_slide.shapes:
            if 'Project review' in shape.text if hasattr(shape, 'text') else False:
                title_shape = shape
                break
        if title_shape and projects:
            project_name = projects[0].get('project', {}).get('name', '')
            set_shape_text(title_shape, f"Project review: {project_name}")

    # Update slide 2 (planning) with milestones if it exists
    if len(prs.slides) > 2:
        planning_slide = prs.slides[2]
        update_planning_slide(planning_slide, projects[0] if projects else {})

    # Save
    prs.save(output_path)
    print(f"✓ PPT saved to: {output_path}")
    print(f"  Slides: {len(prs.slides)}")
    print(f"  Projects: {len(projects)}")

    return output_path


def add_project_review_content(slide, proj_data, reference_data):
    """Add project review content to a new slide."""
    project = proj_data.get('project', {})
    resolved = proj_data.get('resolved', {})
    milestones = proj_data.get('milestones', [])
    attention_points = proj_data.get('attention_points', [])

    # Title
    title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.17), Inches(6.5), Inches(0.43))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = f"Project review : {project.get('name', 'Unknown')}"
    p.font.size = Pt(18)
    p.font.bold = True

    # Date
    date_box = slide.shapes.add_textbox(Inches(8.88), Inches(0.08), Inches(1.07), Inches(0.35))
    tf = date_box.text_frame
    p = tf.paragraphs[0]
    p.text = datetime.now().strftime("%d/%m/%y")
    p.font.size = Pt(10)

    # Status/Mood box (top left)
    status_box = slide.shapes.add_textbox(Inches(0.39), Inches(0.87), Inches(4.53), Inches(0.96))
    tf = status_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    mood = resolved.get('mood', project.get('mood', ''))
    status = resolved.get('status', project.get('status', ''))
    owner = project.get('owner', {})
    owner_name = owner.get('name', '') if isinstance(owner, dict) else ''
    p.text = f"Status: {status}\nMood: {mood}\nOwner: {owner_name}"
    p.font.size = Pt(10)

    # Info box
    info_box = slide.shapes.add_textbox(Inches(0.39), Inches(2.05), Inches(4.53), Inches(0.75))
    tf = info_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    start_date = format_date(project.get('start_date', ''))
    end_date = format_date(project.get('end_date', ''))
    ms_count = len(milestones) if milestones else 0
    completed_ms = len([m for m in milestones if m.get('is_completed')]) if milestones else 0
    p.text = f"Start: {start_date}\nEnd: {end_date}\nMilestones: {completed_ms}/{ms_count}"
    p.font.size = Pt(9)

    # Achievements box
    ach_box = slide.shapes.add_textbox(Inches(0.39), Inches(3.04), Inches(4.53), Inches(1.12))
    tf = ach_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    desc = project.get('description_text', '')
    p.text = desc[:250] + ('...' if len(desc) > 250 else '') if desc else "No description"
    p.font.size = Pt(9)

    # Made box (right top)
    made_box = slide.shapes.add_textbox(Inches(5.14), Inches(0.87), Inches(4.51), Inches(1.19))
    tf = made_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    completed = [m for m in milestones if m.get('is_completed')][:3] if milestones else []
    if completed:
        made_text = "Made:\n" + "\n".join([f"• {m.get('name', '')}" for m in completed])
    else:
        made_text = "Made:\nNo completed milestones"
    p.text = made_text
    p.font.size = Pt(9)

    # Risks box (right middle)
    risks_box = slide.shapes.add_textbox(Inches(5.14), Inches(2.29), Inches(4.41), Inches(1.04))
    tf = risks_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    risk_level = resolved.get('risk', project.get('risk', 'Not set'))
    if attention_points:
        ap_text = "\n".join([f"• {ap.get('title', '')[:40]}" for ap in attention_points[:2]])
        p.text = f"Risk: {risk_level}\n{ap_text}"
    else:
        p.text = f"Risk: {risk_level}\nNo attention points"
    p.font.size = Pt(9)

    # Budget box (right bottom)
    budget_box = slide.shapes.add_textbox(Inches(5.14), Inches(4.16), Inches(4.43), Inches(1.15))
    tf = budget_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Build\nBAC: N/A (API unavailable)\nActual: N/A\nEAC: N/A"
    p.font.size = Pt(9)


def update_planning_slide(slide, proj_data):
    """Update planning slide with milestone info."""
    milestones = proj_data.get('milestones', [])

    # Find the planning text shape
    for shape in slide.shapes:
        if shape.has_text_frame:
            if 'Insert planning' in shape.text:
                if milestones:
                    ms_text = "Milestones:\n" + "\n".join([
                        f"{'✓' if m.get('is_completed') else '○'} {m.get('name', '')} - {m.get('date', '')}"
                        for m in milestones[:8]
                    ])
                else:
                    ms_text = "No milestones defined for this project"
                set_shape_text(shape, ms_text)
                break


def generate_ppt_simple(data_path, output_path):
    """Generate simple PPT without template (fallback)."""
    data = load_data(data_path)

    # Create presentation (16:9)
    prs = Presentation()
    prs.slide_width = Inches(10)  # Match template
    prs.slide_height = Inches(5.625)

    projects = data.get('projects', [])
    reference_data = data.get('reference_data', {})

    # Add slides for each project
    for proj_data in projects:
        project = proj_data.get('project', {})
        resolved = proj_data.get('resolved', {})

        # Add a simple slide
        slide_layout = prs.slide_layouts[6]  # Blank
        slide = prs.slides.add_slide(slide_layout)

        # Title
        title_box = slide.shapes.add_textbox(Inches(0.4), Inches(0.17), Inches(8), Inches(0.5))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"Project Review: {project.get('name', 'Unknown')}"
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

    prs.save(output_path)
    print(f"✓ PPT saved to: {output_path}")
    print(f"  Slides: {len(prs.slides)}")
    print(f"  Projects: {len(projects)}")

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
