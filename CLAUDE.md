# Flash Report Custom AirSaas - Claude Code Instructions

## Specification Reference

**IMPORTANT:** The complete project specification is in `SPEC.md`. Always consult this file for:
- User flows and workflows
- Commands reference
- PPT output structure requirements
- Acceptance criteria
- Known missing fields

---

## Documentation Sync Rule

**IMPORTANT:** When updating any documentation or API-related information, keep ALL documentation files synchronized:

| File | Content |
|------|---------|
| `SPEC.md` | Project specification - source of truth for requirements |
| `CLAUDE.md` | Claude Code instructions, API reference, mapping |
| `README.md` | User documentation, Quick Start, endpoints |
| `scripts/airsaas_fetcher.gs` | Google Apps Script with API endpoints |
| `config/mapping.json` | Field mappings with sources and status |
| `.claude/commands/fetch.md` | Fetch command with endpoints list |
| `.claude/commands/ppt-gamma.md` | Gamma API reference |
| `tracking/MISSING_FIELDS.md` | Available vs missing fields |

When making changes to:
- **API endpoints** → Update: CLAUDE.md, README.md, airsaas_fetcher.gs, fetch.md
- **Field availability** → Update: CLAUDE.md, README.md, mapping.json, MISSING_FIELDS.md, airsaas_fetcher.gs
- **Gamma API** → Update: CLAUDE.md, README.md, ppt-gamma.md

**ALWAYS keep `scripts/airsaas_fetcher.gs` updated** - it serves as the reference implementation for AirSaas API usage and must reflect all discovered endpoints and fields.

---

## Project Overview

This project automates the generation of PowerPoint portfolio reports from AirSaas project data. It fetches project information via the AirSaas API, maps data fields to PPT template placeholders, and generates presentations using either Gamma API or python-pptx.

**Target SmartView:** Projets Vitaux CODIR
- URL: https://app.airsaas.io/space/aqme-corp-/projects/portfolio/projets-vitaux-codir_f67bb94f-464e-44e1-88bc-101bccaed640

**Workspace:** `aqme-corp-`

**Reference PPT Template:** `templates/ProjectCardAndFollowUp.pptx`

---

## Available Commands

| Command | Description | Input | Output |
|---------|-------------|-------|--------|
| `/fetch` | Fetch AirSaas data | config/projects.json | data/{date}_projects.json |
| `/mapping` | Interactive field mapping | Template + user answers | config/mapping.json |
| `/ppt-gamma` | Generate PPT via Gamma API | Fetched data + mapping | outputs/{date}_portfolio_gamma.pptx |
| `/ppt-skill` | Generate PPT via python-pptx | Fetched data + mapping | outputs/{date}_portfolio_skill.pptx |
| `/ppt-all` | Fetch + Generate both | All above | All above |
| `/config` | View/edit configuration | - | Display or prompt |

---

## Project Structure

```
flash-report-custom-airsaas-poc/
├── CLAUDE.md                    # This file - instructions for Claude Code
├── SPEC.md                      # Project specification (source of truth)
├── README.md                    # Project documentation
├── .env                         # Credentials (gitignored)
├── .env.example                 # Template for .env
├── config/
│   ├── projects.json            # Project IDs to export
│   └── mapping.json             # Field mapping AirSaas → PPT
├── templates/
│   └── ProjectCardAndFollowUp.pptx  # Reference PPT template
├── data/
│   └── {date}_projects.json     # Fetched data cache
├── outputs/
│   └── {date}_portfolio_*.pptx  # Generated presentations
├── tracking/
│   ├── MISSING_FIELDS.md        # API fields not available
│   └── CLAUDE_ERRORS.md         # Errors log to avoid repeating
└── scripts/
    ├── generate_ppt.py          # Python script for PPT generation (python-pptx)
    ├── generate_ppt_gamma.py    # Python script for PPT generation (Gamma API)
    └── airsaas_fetcher.gs       # Google Apps Script reference
```

---

## Configuration

### Environment Variables (.env)

| Variable | Description | Example |
|----------|-------------|---------|
| `AIRSAAS_API_KEY` | AirSaas API key | (plain key, no prefix) |
| `AIRSAAS_BASE_URL` | AirSaas API base URL | https://api.airsaas.io/v1 |
| `GAMMA_API_KEY` | Gamma API key | sk-gamma-xxx |
| `GAMMA_BASE_URL` | Gamma API base URL | https://public-api.gamma.app/v1.0 |
| `GAMMA_TEMPLATE_ID` | Gamma template ID | g_9d4wnyvr02om4zk |

### projects.json Structure

```json
{
  "workspace": "aqme-corp-",
  "projects": [
    {
      "id": "UUID-of-project",
      "short_id": "ABC123",
      "name": "Project Name"
    }
  ]
}
```

**Note:** SmartView project lists have NO API endpoint. Project IDs must be added manually by:
1. Going to the SmartView in AirSaas UI
2. Opening each project
3. Copying the UUID from the URL

---

## Systra Template Structure

The template `systra_template.pptx` has:
- **Dimensions:** 10.00 x 5.62 inches
- **3 Slides:**

### Slide 1: Project Review (Main)
| Shape Name | Position | Content |
|------------|----------|---------|
| `Google Shape;671;p1` | Title | "Project review : {name}" |
| `Google Shape;681;p1` | Top right | Date (dd/mm/yy) |
| `Google Shape;678;p1` | Top left | Status, Mood, Owner |
| `Google Shape;677;p1` | Mid left | Start/End dates, Milestones count |
| `Google Shape;673;p1` | Bottom left | Achievements/Description |
| `Google Shape;679;p1` | Bottom left | Next steps |
| `Google Shape;676;p1` | Top right | Made (completed milestones) |
| `Google Shape;680;p1` | Mid right | Risks/Attention points |
| `Google Shape;675;p1` | Bottom right | Budget (BAC/Actual/EAC) |

### Slide 2: Budget Table
- Title placeholder
- Budget tracking table
- Link to budget file

### Slide 3: Planning
- Timeline area for milestones
- Team efforts table (10 columns: Cyber, Infra, Devops, Data, M365, Solution, MDC, PM, Business)

### Available Layouts (10)
| # | Layout Name | Use |
|---|-------------|-----|
| 0 | 2_Ipo_Project_Review_slide_1 | Main project review |
| 1 | Texte + message clef | Text with key message |
| 2 | Ipo_Project_Review_Slide_2 | Planning/Timeline |
| 3 | Texte | Text only |
| 4 | 2 Textes | Two text areas |
| 5 | Texte + soustitre | Text with subtitle |
| 6 | Project_Card | Project card |
| 7 | 1_Ipo_Project_Review_slide_1 | Alt project review |
| 8 | Texte + message clef + soustitre | Text + key + subtitle |
| 9 | Chapitre image vert | Chapter with green image |

---

## AirSaas API Reference

**Documentation:** https://developers.airsaas.io/reference/introduction

### Authentication

```
Authorization: Api-Key {AIRSAAS_API_KEY}
```

**Important:** The header format is exactly `Api-Key ` followed by your key (with a space).

### Base URL

```
https://api.airsaas.io/v1
```

HTTPS is required for all requests.

### Available Endpoints

#### Core Endpoints
| Endpoint | Method | Description | Expandable Fields |
|----------|--------|-------------|-------------------|
| `/projects/` | GET | List all projects | owner, program, goals, teams, requesting_team |
| `/projects/{id}/` | GET | Single project | owner, program, goals, teams, requesting_team |
| `/milestones/` | GET | All milestones | owner, team, project |
| `/decisions/` | GET | All decisions | owner, decision_maker, project, workspace_program |
| `/attention_points/` | GET | All attention points | owner, project, workspace_program |

#### Reference Data
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/projects_moods/` | GET | Mood/weather codes |
| `/projects_statuses/` | GET | Status codes |
| `/projects_risks/` | GET | Risk codes |

#### Project Sub-resources
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/projects/{id}/members/` | GET | Project team members with roles |
| `/projects/{id}/efforts/` | GET | Per-team effort breakdown |

#### Workspace Resources
| Endpoint | Method | Description |
|----------|--------|-------------|
| `/teams/` | GET | All workspace teams |
| `/users/` | GET | Workspace members |
| `/programs/` | GET | Programs list |
| `/project_custom_attributes/` | GET | Custom attribute definitions |
| `/profile/` | GET | Current user profile |

### Expandable Properties

Use `?expand=field1,field2` to include related objects:
- Single: `?expand=owner`
- Multiple: `?expand=owner,program,goals`
- All: `?expand=*`

Without expansion, related fields return only UUIDs.

### Filtering by Project

For milestones, decisions, and attention_points, filter by project:
```
GET /milestones/?project={project_id}&expand=owner,team,project
GET /decisions/?project={project_id}&expand=owner,decision_maker,project
GET /attention_points/?project={project_id}&expand=owner,project
```

### Pagination

| Parameter | Description | Default | Maximum |
|-----------|-------------|---------|---------|
| `page` | Page number | 1 | - |
| `page_size` | Items per page | 10 | **20** |

**Response format:**
```json
{
  "count": 123,
  "next": "https://api.airsaas.io/v1/endpoint/?page=2",
  "previous": null,
  "results": [...]
}
```

**Important:** Always follow `next` link until null to get all results.

### Rate Limits

| Limit | Value |
|-------|-------|
| Per second | 15 calls |
| Per minute | 500 calls |
| Per day | 100,000 calls |

On 429 error:
- Response body: `"Request slowed down. Expected available in X seconds."`
- Header: `Retry-After: {seconds}`

### Date/Time Formats (ISO 8601)

- Datetime: `2023-07-12T18:12:45.123Z`
- Date only: `2023-07-12`

---

## Data Mapping: AirSaas → Template

### Available Fields (from API)

| Template Field | AirSaas Source | Status |
|----------------|----------------|--------|
| Project Name | `project.name` | OK |
| Short ID | `project.short_id` | OK |
| Status | `resolved.status` / `project.status` | OK |
| Mood | `resolved.mood` / `project.mood` | OK |
| Risk | `resolved.risk` / `project.risk` | OK |
| Owner | `project.owner.name` | OK |
| Program | `project.program.name` | OK |
| Start Date | `project.start_date` | OK |
| End Date | `project.end_date` | OK |
| Description | `project.description_text` | OK |
| Teams | `project.teams[].name` | OK |
| Milestones | `milestones[]` | OK |
| Decisions | `decisions[]` | OK |
| Attention Points | `attention_points[]` | OK |
| Budget BAC | `project.budget_capex_initial` | OK |
| Budget Actual | `project.budget_capex_used` | OK |
| Budget EAC | `project.budget_capex_landing` | OK |
| Effort Planned | `project.effort` | OK |
| Effort Used | `project.effort_used` | OK |
| Team Efforts | `/projects/{id}/efforts/` | OK |
| Members | `/projects/{id}/members/` | OK |
| Progress | `project.progress` | OK |
| Milestone Progress | `project.milestone_progress` | OK |
| Gain | `project.gain` | OK |

### Missing Fields (not in API)

| Template Field | Reason | Workaround |
|----------------|--------|------------|
| Mood comment | API returns code only | Use mood label |
| Deployment area | Field not in API | Leave empty |
| End users (actual/target) | Field not in API | Leave empty |

### Reference Data (for code resolution)

```json
{
  "moods": [
    {"code": "blocked", "name_fr": "Bloqué", "color": "#FF0A55"},
    {"code": "complicated", "name_fr": "C'est compliqué", "color": "#FF922B"},
    {"code": "issues", "name_fr": "Quelques problèmes", "color": "#FFD43B"},
    {"code": "good", "name_fr": "Tout va bien", "color": "#03E26B"}
  ],
  "statuses": [
    {"code": "idea", "name": "Idée"},
    {"code": "ongoing", "name": "En cours"},
    {"code": "paused", "name": "En pause"},
    {"code": "finished", "name": "Fini"},
    {"code": "canceled", "name": "Annulé"}
    // ... more statuses
  ],
  "risks": [
    {"code": "low", "name_fr": "Faible"},
    {"code": "medium", "name_fr": "Moyen"},
    {"code": "high", "name_fr": "Élevé"}
  ]
}
```

---

## Gamma API Reference (v1.0)

**Documentation:** https://developers.gamma.app/docs/getting-started

**Template ID:** `g_9d4wnyvr02om4zk` (configured in `.env` as `GAMMA_TEMPLATE_ID`)

### Authentication

```
X-API-KEY: {GAMMA_API_KEY}
```

### Primary Method: Create from Template (Used by /ppt-gamma)

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations/from-template`

**Request:**
```json
{
  "gammaId": "g_9d4wnyvr02om4zk",
  "prompt": "Fill this presentation with project data: ...",
  "exportAs": "pptx",
  "imageOptions": {
    "source": "noImages"
  }
}
```

### Key Parameters (from-template)

| Parameter | Required | Description |
|-----------|----------|-------------|
| `gammaId` | Yes | Template gamma ID (e.g., `g_9d4wnyvr02om4zk`) |
| `prompt` | Yes | Instructions for content modification (max 100,000 tokens) |
| `exportAs` | No | `pdf` or `pptx` for download |
| `themeId` | No | Visual styling override |
| `imageOptions` | No | Image generation settings |

### Alternative Method: Standard Generation

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations`

**Request:**
```json
{
  "inputText": "Your markdown content here",
  "textMode": "preserve",
  "format": "presentation",
  "numCards": 10,
  "cardSplit": "inputTextBreaks",
  "exportAs": "pptx",
  "textOptions": {
    "amount": "medium",
    "language": "fr"
  },
  "cardOptions": {
    "dimensions": "16x9"
  },
  "imageOptions": {
    "source": "noImages"
  }
}
```

### Text Modes (Standard Generation)

| Mode | Use Case |
|------|----------|
| `generate` | Brief input, let AI expand |
| `condense` | Large input, summarize it |
| `preserve` | Keep exact text as-is |

### Additional Options

| Option | Parameters |
|--------|------------|
| `textOptions` | `amount`, `tone`, `audience`, `language` |
| `cardOptions` | `dimensions` (16x9, 4x3, fluid), `headerFooter` |
| `imageOptions` | `source` (noImages, aiGenerated, unsplash...), `model`, `style` |
| `sharingOptions` | `workspaceAccess`, `externalAccess`, `emailOptions` |

### All Gamma Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/v1.0/generations` | POST | Generate from text |
| `/v1.0/generations/from-template` | POST | **Generate from template (primary method)** |
| `/v1.0/themes` | GET | List available themes |
| `/v1.0/folders` | GET | List folders |

### Response

Response includes `pptxUrl` when `exportAs: "pptx"`. Download immediately as URLs expire.

### API Access

- Available on Pro, Ultra, Teams, Business plans
- Credit-based billing
- v0.2 deprecated Jan 16, 2026

**Full reference:** See `.claude/commands/ppt-gamma.md` for complete API documentation.

---

## PPT Generation (python-pptx)

### Script Location
`scripts/generate_ppt.py`

### Script Commands

| Command | Description |
|---------|-------------|
| `python3 scripts/generate_ppt.py` | Generate PPT (with template verification) |
| `python3 scripts/generate_ppt.py --verify` | Verify template compatibility only |
| `python3 scripts/generate_ppt.py --analyze` | Analyze template and show shape positions |

### Template Verification

The script automatically verifies the template before generating:

1. **Checks shape positions** against `EXPECTED_SHAPE_POSITIONS` constant
2. **Reports warnings** if shapes have moved or are missing
3. **Continues generating** even if verification fails (with warning)

```
Verifying template compatibility...
✓ Template verified OK
```

If template changed significantly:
```
⚠️  Template may have changed! Only 60% of expected shapes found.
   Run: python3 scripts/generate_ppt.py --analyze
   Then update EXPECTED_SHAPE_POSITIONS in the script.
```

### Output Structure (per spec section 5.1)

| Slide | Content |
|-------|---------|
| 1 | **Summary** - List of all projects with mood/status |
| 2 to N | **Project slides** - One slide per project with details |
| Last | **Data Notes** - List of unfilled fields and API limitations |

### How It Works

1. **Verifies template** - Checks shape positions match expectations
2. Creates **Summary slide** with all projects listed (ID, Name, Status, Mood, Owner)
3. Reads latest data from `data/{date}_projects.json`
4. Creates **Project slides** - one per project with:
   - Title, date
   - Status, mood, owner
   - Start/end dates, milestones
   - Achievements, next steps
   - Budget, risks
5. Creates **Data Notes slide** listing:
   - Known API limitations (mood comment, deployment area, end users)
   - Per-project unfilled fields
6. Saves to `outputs/{date}_portfolio_skill.pptx`

### Template Handling Rules (CRITICAL)

When working with any PPT template, follow these rules to preserve visual elements:

1. **Analyze template first** - Run python-pptx to extract shape positions and sizes
2. **Duplicate slides via XML** - Preserves images, borders, colors, fonts
3. **Find shapes by position** - Names change after duplication; use Euclidean distance
4. **Delete placeholders entirely** - On Summary/Data Notes slides (not just clear text)
5. **Handle visual section titles** - Add leading newline if template has titles in background
6. **Handle overlapping shapes** - Some shapes physically overlap; use only one and clear the other

### Shape Overlap Handling (CRITICAL)

The ProjectCardAndFollowUp.pptx template has overlapping shapes:

| Overlapping Shapes | Action |
|--------------------|--------|
| 672 + 677 | Use 672 for SCOPE content, CLEAR 677 |
| 674 + 679 | Use 679 for NEXT STEPS, CLEAR 674 |

**Layout Title Overlap (Shape 672):**
- Layout title "SCOPE & MILESTONES" (Y=1.83-2.09) overlaps shape 672 (Y=1.93)
- Solution: Use `set_shape_text_with_title_padding()` with `padding_lines=1`
- First paragraph MUST be empty to push content below title

```python
# Shape with overlapping layout title
set_shape_text_with_title_padding(
    shape,
    items=["Item 1", "Item 2"],
    padding_lines=1,  # Empty first line
    font_size=9,
    use_bullets=True
)

# Overlapping shape - CLEAR IT
set_shape_text(overlapping_shape, '', font_size=9)
```

### Template Color Palette

Use these colors from `mapping.json._color_palette`:

| Element | Hex Code | Use |
|---------|----------|-----|
| Section titles | #003D4B | Layout titles (dark teal) |
| Content text | #000000 | All body text (black) |
| Mood: good | #03E26B | Green |
| Mood: issues | #FFD43B | Yellow |
| Mood: complicated | #FF922B | Orange |
| Mood: blocked | #FF0A55 | Red |
| Risk: low/medium/high | Same as moods | Green/Yellow/Red |

```python
# Analyze current template
from pptx import Presentation
prs = Presentation('templates/YOUR_TEMPLATE.pptx')
for slide in prs.slides:
    for shape in slide.shapes:
        x, y = shape.left.inches, shape.top.inches
        print(f"{shape.name}: ({x:.2f}, {y:.2f})")
```

### Font Size Guidelines

| Element | Recommended Size |
|---------|------------------|
| Slide title | 14-20pt |
| Section headers | 10-12pt |
| Body content | 8-10pt |
| Table content | 7-8pt |
| Date/metadata | 8pt |
| Dense notes | 6-7pt |

**Always analyze template and adjust font sizes to fit content in shapes.**

---

## Template Style Extraction System

**CRITICAL:** The script MUST extract visual styles from the template at runtime and apply them uniformly. This ensures visual consistency when templates change.

### Configuration Location

Style extraction rules are defined in `config/mapping.json` under `_style_extraction_rules`:

```json
{
  "_style_extraction_rules": {
    "style_types": {
      "title": {"extract_from_position": {"x": 0.39, "y": 0.17}},
      "bullet": {"skip_colors": ["000000", "FFFFFF"]}
    },
    "defaults": {
      "font_name": "Lato",
      "bullet_color": "D32427"
    }
  }
}
```

### Style Types

| Style | Source Shape | Applied To |
|-------|--------------|------------|
| `title` | Shape 671 (0.39, 0.17) | Slide titles |
| `date` | Shape 681 (8.88, 0.08) | Date fields |
| `status` | Shape 678 (0.39, 0.98) | Status/mood text |
| `inline_title` | Shapes 676, 675 | "Build", "Made :" |
| `content` | Shapes 672, 673, 679, 680 | Bullet items |
| `bullet` | First colored bullet found | ALL bullet points |

### Extraction Process

1. Load rules from `mapping.json._style_extraction_rules`
2. Scan template slide 0 shapes
3. Extract font name, size, color, bold from each style type position
4. Extract bullet character and color (skip black/white)
5. Store in `TEMPLATE_STYLES` global
6. Apply defaults from mapping.json if extraction fails

### Script Functions

| Function | Purpose |
|----------|---------|
| `load_style_extraction_rules()` | Read rules from mapping.json |
| `extract_all_template_styles(slide)` | Extract styles from template |
| `apply_template_style(para, style_type)` | Apply extracted style to paragraph |
| `_add_bullet(para)` | Add bullet with template color |

### Application Rules

From `mapping.json._style_extraction_rules.application_rules`:

1. **NEVER hardcode** font names, sizes, or colors in script
2. **ALWAYS extract** from template at runtime
3. **Bullet color MUST be uniform** across all slides
4. **Font family MUST match** template exactly
5. When template changes, styles re-extract automatically

### Updating for New Templates

When the template changes:

1. Run `/mapping` to re-analyze shapes
2. Update `_style_extraction_rules.style_types` positions if needed
3. Update `_style_extraction_rules.defaults` if defaults should change
4. Script will automatically extract new styles on next run

### Dependencies

```bash
pip install python-pptx
```

---

## Error Handling

### Logging Errors

Log all errors to `tracking/CLAUDE_ERRORS.md` with:
- Date and time
- Command that failed
- Error message
- Solution applied
- Prevention for future

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| 401 Unauthorized | Invalid or missing API key | Check `.env` AIRSAAS_API_KEY |
| 404 Not Found | Invalid project ID or endpoint | Verify UUID format |
| 429 Too Many Requests | Rate limit exceeded | Wait `Retry-After` seconds |
| Empty results | Wrong filter or no data | Check project ID filter |
| `ImportError: RgbColor` | Wrong import in python-pptx | Use `from pptx.dml.color import RGBColor` |

---

## Workflow Reference

### Initial Setup (One-time)

1. Create `.env` from `.env.example`
2. Add API keys (AirSaas, optionally Gamma)
3. Add project IDs to `config/projects.json`
4. Place template in `templates/systra_template.pptx`
5. Run `/mapping` to configure field mapping (optional)

### Regular Usage

```
/fetch          → Fetches data, saves to data/{date}_projects.json
/ppt-skill      → Generates PPT via python-pptx using template
/ppt-gamma      → Generates PPT via Gamma API (requires Gamma key)
```

Or all at once:
```
/ppt-all        → Fetch + generate both versions
```

### Maintenance

```
/config         → View current configuration
/config projects → Edit project list
/mapping        → Re-run if template changes (syncs EXPECTED_SHAPE_POSITIONS)
```

### Template Changes Workflow

When the PPT template is modified:

1. **Run `/mapping`** - This will:
   - Verify current template compatibility
   - Analyze new shape positions
   - Update `EXPECTED_SHAPE_POSITIONS` in `generate_ppt.py`
   - Update field mappings if needed

2. **Or manually**:
   ```bash
   python3 scripts/generate_ppt.py --analyze   # See new positions
   python3 scripts/generate_ppt.py --verify    # Test compatibility
   ```

3. **Update script** if positions changed significantly:
   - Edit `EXPECTED_SHAPE_POSITIONS` constant
   - Edit position arguments in `populate_project_slide()` function

---

## Development Notes

- The API uses RESTful principles with JSON request/response bodies
- All temporal values follow ISO 8601 format
- SmartView project lists have NO API endpoint - manual ID entry required
- Always save fetched data with date prefix for caching/history
- Template analysis happens during `/mapping` phase
- Both Gamma (AI-styled) and python-pptx (exact control) outputs available
- python-pptx requires: `pip install python-pptx`
- Template shapes are identified by exact name (e.g., `Google Shape;671;p1`)
