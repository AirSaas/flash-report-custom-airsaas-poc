#!/usr/bin/env python3
"""
PPT Generator for AirSaas Flash Reports via Gamma API
Uses Gamma's from-template API to generate PowerPoint presentations.
"""

import json
import os
import sys
import requests
import time
from datetime import datetime
from glob import glob

# =============================================================================
# CONFIGURATION
# =============================================================================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(BASE_DIR, 'data')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
CONFIG_DIR = os.path.join(BASE_DIR, 'config')

# Default polling settings
POLL_INTERVAL_SECONDS = 5
MAX_POLL_ATTEMPTS = 60  # 5 minutes max


# =============================================================================
# ENVIRONMENT LOADING
# =============================================================================

def load_env(filepath=None):
    """Load environment variables from .env file without external dependencies."""
    if filepath is None:
        filepath = os.path.join(BASE_DIR, '.env')

    env_vars = {}
    if not os.path.exists(filepath):
        print(f"Warning: .env file not found at {filepath}")
        return env_vars

    with open(filepath, 'r') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#') and '=' in line:
                key, value = line.split('=', 1)
                env_vars[key.strip()] = value.strip()
    return env_vars


# =============================================================================
# DATA LOADING
# =============================================================================

def find_latest_data_file():
    """Find the most recent projects JSON file in data directory."""
    pattern = os.path.join(DATA_DIR, '*_projects.json')
    files = glob(pattern)

    if not files:
        return None

    # Sort by modification time, newest first
    files.sort(key=os.path.getmtime, reverse=True)
    return files[0]


def load_project_data(filepath=None):
    """Load project data from JSON file."""
    if filepath is None:
        filepath = find_latest_data_file()

    if filepath is None:
        raise FileNotFoundError("No project data found. Run /fetch first.")

    print(f"Loading data from: {filepath}")

    with open(filepath, 'r') as f:
        return json.load(f)


def load_mapping():
    """Load field mapping configuration."""
    mapping_path = os.path.join(CONFIG_DIR, 'mapping.json')

    if os.path.exists(mapping_path):
        with open(mapping_path, 'r') as f:
            return json.load(f)

    return {}


# =============================================================================
# REFERENCE DATA RESOLUTION
# =============================================================================

def build_reference_lookups(data):
    """Build lookup tables for status, mood, risk codes."""
    ref_data = data.get('reference_data', {})

    moods = {m['code']: m for m in ref_data.get('moods', [])}
    statuses = {s['code']: s for s in ref_data.get('statuses', [])}
    risks = {r['code']: r for r in ref_data.get('risks', [])}

    return moods, statuses, risks


def resolve_mood(code, moods):
    """Resolve mood code to label."""
    if code and code in moods:
        return moods[code].get('name_fr', moods[code].get('name', code))
    return code or 'N/A'


def resolve_status(code, statuses):
    """Resolve status code to label."""
    if code and code in statuses:
        return statuses[code].get('name', code)
    return code or 'N/A'


def resolve_risk(code, risks):
    """Resolve risk code to label."""
    if code and code in risks:
        return risks[code].get('name_fr', risks[code].get('name', code))
    return code or 'N/A'


# =============================================================================
# CONTENT FORMATTING
# =============================================================================

def format_summary_table(projects, moods, statuses):
    """Build markdown summary table of all projects."""
    lines = []

    for p in projects:
        proj = p.get('project', {})
        name = proj.get('name', 'Unknown')
        short_id = proj.get('short_id', 'N/A')
        status = resolve_status(proj.get('status'), statuses)
        mood = resolve_mood(proj.get('mood'), moods)

        owner_data = proj.get('owner')
        owner = owner_data.get('name', 'N/A') if isinstance(owner_data, dict) else 'N/A'

        # Truncate name if too long
        if len(name) > 40:
            name = name[:37] + '...'

        lines.append(f"| {short_id} | {name} | {status} | {mood} | {owner} |")

    header = "| ID | Project | Status | Mood | Owner |\n|---|---|---|---|---|"
    return header + "\n" + "\n".join(lines)


def format_project_section(p, moods, statuses, risks):
    """Format a single project as markdown section."""
    proj = p.get('project', {})

    # Basic info
    name = proj.get('name', 'Unknown')
    short_id = proj.get('short_id', 'N/A')
    status = resolve_status(proj.get('status'), statuses)
    mood = resolve_mood(proj.get('mood'), moods)
    risk = resolve_risk(proj.get('risk'), risks)

    # Owner & Program
    owner_data = proj.get('owner')
    owner = owner_data.get('name', 'N/A') if isinstance(owner_data, dict) else 'N/A'

    program_data = proj.get('program')
    program = program_data.get('name', 'N/A') if isinstance(program_data, dict) else 'N/A'

    # Dates
    start_date = proj.get('start_date', 'N/A')
    end_date = proj.get('end_date', 'N/A')

    # Budget
    bac = proj.get('budget_capex_initial') or 0
    actual = proj.get('budget_capex_used') or 0
    eac = proj.get('budget_capex_landing') or 0

    # Effort
    effort = proj.get('effort') or 0
    effort_used = proj.get('effort_used') or 0

    # Progress
    progress = proj.get('progress') or 0

    # Description
    description = proj.get('description_text', '') or ''
    if len(description) > 200:
        description = description[:200] + '...'

    # Milestones
    milestones = p.get('milestones', [])
    completed_milestones = [m for m in milestones if m.get('status') == 'done']
    total_milestones = len(milestones)
    completed_count = len(completed_milestones)

    # Get 3 recent completed milestones
    made_items = []
    for m in completed_milestones[:3]:
        made_items.append(f"- {m.get('name', 'Milestone')}")

    # Decisions (pending)
    decisions = p.get('decisions', [])
    pending_decisions = [d for d in decisions if d.get('status') not in ['taken', 'actions-done']]
    next_steps = []
    for d in pending_decisions[:3]:
        next_steps.append(f"- {d.get('name', 'Decision')}")

    # Attention points
    attention_points = p.get('attention_points', [])
    risks_list = []
    for ap in attention_points[:3]:
        risks_list.append(f"- {ap.get('name', 'Attention point')}")

    # Build section
    section = f"""
---
## Project: {name} ({short_id})

**Status:** {status} | **Mood:** {mood} | **Risk:** {risk}

**Owner:** {owner}
**Program:** {program}

### Key Information
- **Timeline:** {start_date} to {end_date}
- **Progress:** {progress}%
- **Milestones:** {completed_count}/{total_milestones} completed

### Effort
- Planned: {effort} days
- Used: {effort_used} days

### Description
{description if description else 'No description available'}

### Made (Completed)
{chr(10).join(made_items) if made_items else '- No completed milestones'}

### Next Steps
{chr(10).join(next_steps) if next_steps else '- No pending decisions'}

### Budget (Build)
- **BAC:** {bac/1000:.0f} K€
- **Actual:** {actual/1000:.0f} K€
- **EAC:** {eac/1000:.0f} K€

### Risks
Risk Level: {risk}
{chr(10).join(risks_list) if risks_list else ''}
"""
    return section


def build_prompt_content(data, mapping=None):
    """Build the full prompt content for Gamma API."""
    projects = data.get('projects', [])
    moods, statuses, risks = build_reference_lookups(data)

    current_date = datetime.now().strftime('%d/%m/%Y')

    # Summary table
    summary_table = format_summary_table(projects, moods, statuses)

    # Project sections
    project_sections = []
    unfilled_fields = []

    for p in projects:
        proj = p.get('project', {})
        short_id = proj.get('short_id', 'N/A')

        # Track unfilled fields
        if not proj.get('mood'):
            unfilled_fields.append(f"- {short_id}: Mood")
        if not proj.get('description_text'):
            unfilled_fields.append(f"- {short_id}: Description")

        milestones = p.get('milestones', [])
        completed = [m for m in milestones if m.get('status') == 'done']
        if not completed:
            unfilled_fields.append(f"- {short_id}: Completed milestones")

        section = format_project_section(p, moods, statuses, risks)
        project_sections.append(section)

    # Build full prompt
    prompt = f"""Fill this presentation template with the following project portfolio data:

# Portfolio Flash Report
**Date:** {current_date}

## Summary

{summary_table}

## Projects
{chr(10).join(project_sections)}

---
## Data Notes

The following fields could not be populated from the API:
- Mood comment (API returns code only)
- Deployment area (Field not available)
- End users actual/target (Field not available)

Unfilled fields:
{chr(10).join(unfilled_fields) if unfilled_fields else '- All fields populated'}
"""

    return prompt, len(projects), unfilled_fields


# =============================================================================
# GAMMA API
# =============================================================================

def call_gamma_api(prompt, api_key, base_url, template_id):
    """Call Gamma API to generate presentation from template."""
    headers = {
        'X-API-KEY': api_key,
        'Content-Type': 'application/json'
    }

    payload = {
        "gammaId": template_id,
        "prompt": prompt,
        "exportAs": "pptx"
    }

    print(f"Calling Gamma API...")
    print(f"  Template ID: {template_id}")
    print(f"  Prompt length: {len(prompt)} chars")

    response = requests.post(
        f"{base_url}/generations/from-template",
        headers=headers,
        json=payload
    )

    if response.status_code in [200, 201]:
        result = response.json()
        generation_id = result.get('generationId')
        print(f"  Generation started: {generation_id}")
        return generation_id, None
    else:
        error_msg = f"HTTP {response.status_code}: {response.text}"
        print(f"  Error: {error_msg}")
        return None, error_msg


def poll_generation(generation_id, api_key, base_url):
    """Poll for generation completion."""
    headers = {
        'X-API-KEY': api_key,
        'Content-Type': 'application/json'
    }

    print(f"\nPolling for completion...")

    for attempt in range(1, MAX_POLL_ATTEMPTS + 1):
        response = requests.get(
            f"{base_url}/generations/{generation_id}",
            headers=headers
        )

        if response.status_code == 200:
            result = response.json()
            status = result.get('status')

            if status == 'completed':
                print(f"  Completed after {attempt} attempts")
                return result, None
            elif status == 'failed':
                error = result.get('error', 'Unknown error')
                return None, f"Generation failed: {error}"
            else:
                print(f"  Attempt {attempt}: {status}")
                time.sleep(POLL_INTERVAL_SECONDS)
        else:
            print(f"  Attempt {attempt}: HTTP {response.status_code}")
            time.sleep(POLL_INTERVAL_SECONDS)

    return None, f"Timeout after {MAX_POLL_ATTEMPTS * POLL_INTERVAL_SECONDS} seconds"


def download_pptx(export_url, output_path):
    """Download PPTX from export URL."""
    print(f"\nDownloading PPTX...")

    response = requests.get(export_url)

    if response.status_code == 200:
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        with open(output_path, 'wb') as f:
            f.write(response.content)

        file_size = len(response.content) / 1024
        print(f"  Saved: {output_path}")
        print(f"  Size: {file_size:.1f} KB")
        return True, None
    else:
        return False, f"Download failed: HTTP {response.status_code}"


# =============================================================================
# MAIN
# =============================================================================

def main():
    """Main entry point."""
    print("=" * 60)
    print("Gamma PPT Generator for AirSaas Flash Reports")
    print("=" * 60)

    # Load environment
    env = load_env()

    GAMMA_API_KEY = env.get('GAMMA_API_KEY')
    GAMMA_BASE_URL = env.get('GAMMA_BASE_URL', 'https://public-api.gamma.app/v1.0')
    GAMMA_TEMPLATE_ID = env.get('GAMMA_TEMPLATE_ID', 'g_9d4wnyvr02om4zk')

    if not GAMMA_API_KEY:
        print("\nError: GAMMA_API_KEY not found in .env file")
        print("Please add your Gamma API key to .env:")
        print("  GAMMA_API_KEY=sk-gamma-xxx")
        sys.exit(1)

    print(f"\nConfiguration:")
    print(f"  API Key: {GAMMA_API_KEY[:15]}...")
    print(f"  Base URL: {GAMMA_BASE_URL}")
    print(f"  Template ID: {GAMMA_TEMPLATE_ID}")

    # Load project data
    try:
        data = load_project_data()
    except FileNotFoundError as e:
        print(f"\nError: {e}")
        print("Run '/fetch' to fetch project data first.")
        sys.exit(1)

    # Load mapping
    mapping = load_mapping()

    # Build prompt content
    print(f"\nBuilding prompt content...")
    prompt, project_count, unfilled = build_prompt_content(data, mapping)
    print(f"  Projects: {project_count}")
    print(f"  Unfilled fields: {len(unfilled)}")

    # Call Gamma API
    generation_id, error = call_gamma_api(
        prompt, GAMMA_API_KEY, GAMMA_BASE_URL, GAMMA_TEMPLATE_ID
    )

    if error:
        print(f"\nFailed to start generation: {error}")
        sys.exit(1)

    # Poll for completion
    result, error = poll_generation(generation_id, GAMMA_API_KEY, GAMMA_BASE_URL)

    if error:
        print(f"\n{error}")
        sys.exit(1)

    # Get URLs and credits
    export_url = result.get('exportUrl')
    gamma_url = result.get('gammaUrl')
    credits = result.get('credits', {})

    print(f"\nGeneration completed:")
    print(f"  Gamma URL: {gamma_url}")
    print(f"  Credits used: {credits.get('deducted', 'N/A')}")
    print(f"  Credits remaining: {credits.get('remaining', 'N/A')}")

    # Download PPTX
    if export_url:
        output_date = datetime.now().strftime('%Y-%m-%d')
        output_path = os.path.join(OUTPUT_DIR, f'{output_date}_portfolio_gamma.pptx')

        success, error = download_pptx(export_url, output_path)

        if not success:
            print(f"\n{error}")
            sys.exit(1)
    else:
        print("\nWarning: No export URL in response")

    # Summary
    print("\n" + "=" * 60)
    print("PPT Generated Successfully!")
    print("=" * 60)
    print(f"\nOutput: {output_path}")
    print(f"Projects: {project_count}")
    print(f"Gamma URL: {gamma_url}")

    if unfilled:
        print(f"\nUnfilled Fields ({len(unfilled)}):")
        for field in unfilled[:10]:
            print(f"  {field}")
        if len(unfilled) > 10:
            print(f"  ... and {len(unfilled) - 10} more")

    print("\nDone!")
    return 0


if __name__ == '__main__':
    sys.exit(main())
