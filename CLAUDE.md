# Flash Report Custom AirSaas - Claude Code Instructions

## Project Overview

This project automates the generation of PowerPoint portfolio reports from AirSaas project data. It fetches project information via the AirSaas API, maps data fields to PPT template placeholders, and generates presentations using either Gamma API or python-pptx.

**Target SmartView:** Projets Vitaux CODIR
- URL: https://app.airsaas.io/space/aqme-corp-/projects/portfolio/projets-vitaux-codir_f67bb94f-464e-44e1-88bc-101bccaed640

**Workspace:** `aqme-corp-`

**Reference PPT Template:** `templates/systra_template.pptx`

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
├── README.md                    # Project documentation
├── .env                         # Credentials (gitignored)
├── .env.example                 # Template for .env
├── config/
│   ├── projects.json            # Project IDs to export
│   └── mapping.json             # Field mapping AirSaas → PPT
├── templates/
│   └── systra_template.pptx     # Reference PPT template (3 slides)
├── data/
│   └── {date}_projects.json     # Fetched data cache
├── outputs/
│   └── {date}_portfolio_*.pptx  # Generated presentations
├── tracking/
│   ├── MISSING_FIELDS.md        # API fields not available
│   └── CLAUDE_ERRORS.md         # Errors log to avoid repeating
└── scripts/
    ├── generate_ppt.py          # Python script for PPT generation
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

### projects.json Structure

```json
{
  "workspace": "aqme-corp-",
  "smartview": {
    "name": "Projets Vitaux CODIR",
    "id": "f67bb94f-464e-44e1-88bc-101bccaed640",
    "url": "https://app.airsaas.io/space/aqme-corp-/projects/portfolio/..."
  },
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

## Gamma API Reference

**Base URL:** `https://public-api.gamma.app/v1.0`

### Authentication

```
X-API-KEY: {GAMMA_API_KEY}
```

### Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/generations` | POST | Create new presentation |
| `/generations/{id}` | GET | Check status / get download URL |

### Creating a Presentation

**Request:**
```json
POST /generations
{
  "input": "Markdown content with \\n---\\n as slide breaks",
  "exportOptions": {
    "pptx": true
  }
}
```

**Slide breaks:** Use `\n---\n` to separate slides in markdown.

### Polling for Completion

```
GET /generations/{id}
```

Poll every 5 seconds until `status: "completed"`, then get download URL from response.

---

## PPT Generation (python-pptx)

### Script Location
`scripts/generate_ppt.py`

### How It Works

1. Loads template from `templates/systra_template.pptx`
2. Reads latest data from `data/{date}_projects.json`
3. Populates slide 1 shapes by name mapping
4. Updates slide 2 (budget) title
5. Updates slide 3 (planning) with milestones
6. Saves to `outputs/{date}_portfolio_skill.pptx`

### Running the Script

```bash
python3 scripts/generate_ppt.py
```

### Shape Name Mapping

The script uses exact shape names from the template:
```python
SLIDE1_SHAPES = {
    'title': 'Google Shape;671;p1',
    'mood_comment': 'Google Shape;678;p1',
    'info': 'Google Shape;677;p1',
    'achievements': 'Google Shape;673;p1',
    'next_steps': 'Google Shape;679;p1',
    'budget': 'Google Shape;675;p1',
    'made': 'Google Shape;676;p1',
    'risks': 'Google Shape;680;p1',
    'date': 'Google Shape;681;p1',
}
```

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
/mapping        → Re-run if template changes
```

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
