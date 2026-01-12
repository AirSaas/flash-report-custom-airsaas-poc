# Configuration Management

View and edit project configuration.

## Usage

- `/config` or `/config show` - Display current configuration
- `/config projects` - Edit project list
- `/config mapping` - View mapping summary (use `/mapping` to edit)
- `/config env` - Check environment variables status

## Instructions

### /config show (default)

Display current configuration status:

```
========================================
Configuration Status
========================================

Environment Variables (.env):
- AIRSAAS_API_KEY: {configured/not set}
- AIRSAAS_BASE_URL: {value or default}
- GAMMA_API_KEY: {configured/not set}
- GAMMA_BASE_URL: {value or default}

Workspace: {workspace from projects.json}

Projects Configured: {count}
{list project names}

Mapping Status:
- Fields mapped: {count with status "ok"}
- Fields missing: {count with status "missing"}
- Fields manual: {count with status "manual"}

Template: {filename if exists, or "not found"}

Last Data Fetch: {date from most recent data file, or "never"}
========================================
```

### /config projects

Interactive editing of project list:

1. **Show Current Projects**
   ```
   Current projects in config/projects.json:
   1. {name} (ID: {short_id})
   2. {name} (ID: {short_id})
   ...
   ```

2. **Offer Options**
   ```
   Options:
   A) Add a project
   B) Remove a project
   C) Replace entire list
   D) Done
   ```

3. **Add Project**
   - Ask for project ID (UUID or short_id)
   - Ask for project name
   - Add to projects.json

4. **Remove Project**
   - Show numbered list
   - Ask which to remove
   - Update projects.json

5. **Replace List**
   - Ask for new list (can paste multiple)
   - Parse and validate
   - Replace in projects.json

### /config mapping

Display mapping summary (read-only):

```
Mapping Summary (config/mapping.json):

Summary Slide:
- title: static "Portfolio Flash Report"
- date: generated (current_date)
- project_list: from projects

Project Card:
- name: project.name [ok]
- status: project.status [ok]
- mood: project.mood [ok]
- mood_comment: [missing]
- budget: budget_values [ok]
- eac_budget: [missing]
...

To modify mappings, run /mapping
```

### /config env

Check environment setup:

```
Environment Check:

AIRSAAS_API_KEY:
- Status: {set/not set}
- Format: {valid/invalid} (should not include "Api-Key" prefix)

AIRSAAS_BASE_URL:
- Status: {set/using default}
- Value: {url}

GAMMA_API_KEY:
- Status: {set/not set}
- Format: {valid/invalid} (should start with "sk-gamma-")

GAMMA_BASE_URL:
- Status: {set/using default}
- Value: {url}

{If any not set, show instructions:}
To configure, edit .env file:
  AIRSAAS_API_KEY=your-key-here
  GAMMA_API_KEY=sk-gamma-your-key-here
```

## File Locations

- Environment: `.env` (gitignored)
- Projects: `config/projects.json`
- Mapping: `config/mapping.json`
- Template: `templates/*.pptx`

## Notes

- Never display actual API key values (show "configured" instead)
- Validate project IDs look like UUIDs when adding
- Warn if projects.json has placeholder values
