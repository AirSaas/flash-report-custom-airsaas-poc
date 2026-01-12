# Flash Report Custom AirSaas - POC

Automates the generation of PowerPoint portfolio reports from AirSaas project data. Fetches project information via the AirSaas API, maps data fields to PPT template placeholders, and generates presentations using python-pptx or Gamma API.

## Features

- Fetch project data from AirSaas API (projects, milestones, decisions, attention points)
- Resolve status/mood/risk codes to human-readable labels
- Generate PowerPoint presentations using a reference template
- Support for multiple output methods: python-pptx (programmatic) and Gamma API (AI-powered)
- Claude Code slash commands for streamlined workflow

## Requirements

- Python 3.8+
- python-pptx library
- AirSaas API key
- (Optional) Gamma API key for AI-powered generation

## Installation

1. Clone this repository
2. Install dependencies:
   ```bash
   pip install python-pptx
   ```
3. Copy `.env.example` to `.env` and add your API keys:
   ```bash
   cp .env.example .env
   ```
4. Place your PPT template in `templates/` folder

## Configuration

### Environment Variables (.env)

| Variable | Description |
|----------|-------------|
| `AIRSAAS_API_KEY` | Your AirSaas API key |
| `AIRSAAS_BASE_URL` | API base URL (default: https://api.airsaas.io/v1) |
| `GAMMA_API_KEY` | Gamma API key (optional) |
| `GAMMA_BASE_URL` | Gamma API URL (default: https://public-api.gamma.app/v1.0) |

### Project Configuration (config/projects.json)

Add the projects you want to export:

```json
{
  "workspace": "your-workspace",
  "smartview": {
    "name": "Your SmartView Name",
    "id": "smartview-uuid"
  },
  "projects": [
    {
      "id": "project-uuid",
      "short_id": "PRJ-001",
      "name": "Project Name"
    }
  ]
}
```

Note: SmartView project lists are not available via API. Project IDs must be added manually from the AirSaas UI.

## Usage

### With Claude Code

Use the following slash commands:

| Command | Description |
|---------|-------------|
| `/fetch` | Fetch data from AirSaas API |
| `/ppt-skill` | Generate PPT using python-pptx |
| `/ppt-gamma` | Generate PPT using Gamma API |
| `/ppt-all` | Fetch data and generate both PPT versions |
| `/mapping` | Configure field mapping interactively |
| `/config` | View or edit configuration |

### Manual Execution

```bash
# Generate PPT from fetched data
python3 scripts/generate_ppt.py
```

## Project Structure

```
flash-report-custom-airsaas-poc/
├── CLAUDE.md                    # Claude Code instructions
├── README.md                    # This file
├── .env                         # API credentials (gitignored)
├── .env.example                 # Credentials template
├── config/
│   ├── projects.json            # Projects to export
│   └── mapping.json             # Field mapping configuration
├── templates/
│   └── *.pptx                   # PPT templates
├── data/
│   └── {date}_projects.json     # Fetched data cache
├── outputs/
│   └── {date}_portfolio_*.pptx  # Generated presentations
├── tracking/
│   ├── MISSING_FIELDS.md        # API fields not available
│   └── CLAUDE_ERRORS.md         # Error log
├── scripts/
│   ├── generate_ppt.py          # Python PPT generator
│   └── airsaas_fetcher.gs       # Google Apps Script reference
└── .claude/commands/            # Slash command definitions
```

## API Reference

### AirSaas API

- **Authentication:** `Authorization: Api-Key {YOUR_KEY}`
- **Base URL:** `https://api.airsaas.io/v1`
- **Pagination:** max `page_size=20`
- **Rate Limits:** 15/sec, 500/min, 100k/day

### Available Endpoints

| Endpoint | Description |
|----------|-------------|
| `/projects/` | List all projects |
| `/projects/{id}/` | Single project details |
| `/milestones/?project={id}` | Project milestones |
| `/decisions/?project={id}` | Project decisions |
| `/attention_points/?project={id}` | Project attention points |
| `/projects_moods/` | Mood code definitions |
| `/projects_statuses/` | Status code definitions |
| `/projects_risks/` | Risk code definitions |

### Expandable Fields

Use `?expand=field1,field2` to include related objects:
- Projects: `owner`, `program`, `goals`, `teams`, `requesting_team`
- Milestones: `owner`, `team`, `project`
- Decisions: `owner`, `decision_maker`, `project`

## Template Mapping

The script maps AirSaas data to template shapes by name. Key mappings:

| Template Field | AirSaas Source |
|----------------|----------------|
| Project Name | `project.name` |
| Status | `resolved.status` |
| Mood | `resolved.mood` |
| Owner | `project.owner.name` |
| Start/End Date | `project.start_date`, `project.end_date` |
| Milestones | `milestones[]` |
| Decisions | `decisions[]` |
| Attention Points | `attention_points[]` |

### Known Missing Fields

These fields are not available in the public API:
- Budget (BAC, Actual, EAC)
- Team efforts
- Deployment area
- End user counts

## License

Internal use only.
