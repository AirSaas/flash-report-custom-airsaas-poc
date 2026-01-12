# Missing Fields Log

This document tracks AirSaas API fields needed for the PPT template.

## Fields Status Update (January 2026)

### ALL REQUIRED FIELDS NOW AVAILABLE

| Field | API Endpoint | Notes |
|-------|--------------|-------|
| Budget CAPEX | `project.budget_capex` | Direct field |
| Budget CAPEX Initial | `project.budget_capex_initial` | BAC equivalent |
| Budget CAPEX Landing | `project.budget_capex_landing` | EAC equivalent |
| Budget CAPEX Used | `project.budget_capex_used` | Actual spent |
| Budget OPEX | `project.budget_opex` | Direct field |
| Budget OPEX Initial | `project.budget_opex_initial` | BAC equivalent |
| Budget OPEX Landing | `project.budget_opex_landing` | EAC equivalent |
| Budget OPEX Used | `project.budget_opex_used` | Actual spent |
| Effort | `project.effort` | Planned effort |
| Effort Used | `project.effort_used` | Actual effort |
| Team Efforts | `/projects/{id}/efforts/` | Per-team breakdown |
| Project Members | `/projects/{id}/members/` | With roles |
| Custom Attributes | `/project_custom_attributes/` | Type, etc. |
| Progress | `project.progress` | 0-100 percentage |
| Milestone Progress | `project.milestone_progress` | 0-100 percentage |
| Gain | `project.gain` | Expected gain |

### STILL MISSING (not in API)

| Field | PPT Location | Workaround |
|-------|--------------|------------|
| Mood comment | Project Card | Use mood label from reference data |
| Deployment area | Project Card | Leave empty or use custom attribute |
| End users (actual) | Project Progress | Leave empty |
| End users (target) | Project Progress | Leave empty |

## Complete API Endpoints

### Core Endpoints
| Endpoint | Description | Expandable Fields |
|----------|-------------|-------------------|
| `GET /projects/` | List projects | owner, program, goals, teams, requesting_team |
| `GET /projects/{id}/` | Single project | owner, program, goals, teams, requesting_team |
| `GET /milestones/?project={id}` | Project milestones | owner, team, project |
| `GET /decisions/?project={id}` | Project decisions | owner, decision_maker, project |
| `GET /attention_points/?project={id}` | Attention points | owner, project |

### Reference Data
| Endpoint | Description |
|----------|-------------|
| `GET /projects_moods/` | Mood definitions |
| `GET /projects_statuses/` | Status definitions |
| `GET /projects_risks/` | Risk definitions |

### Project Sub-resources
| Endpoint | Description |
|----------|-------------|
| `GET /projects/{id}/members/` | Project team members with roles |
| `GET /projects/{id}/efforts/` | Per-team effort breakdown |

### Workspace Resources
| Endpoint | Description |
|----------|-------------|
| `GET /teams/` | All teams |
| `GET /users/` | Workspace members |
| `GET /programs/` | Programs |
| `GET /project_custom_attributes/` | Custom attribute definitions |

## Project Fields (Full List)

From `GET /projects/?expand=*`:

```json
{
  "id": "uuid",
  "name": "string",
  "short_id": "string",
  "owner": { "id", "name", "initials", ... },
  "status": "string (code)",
  "url": "string",
  "archived": "boolean",
  "budget_capex": "number|null",
  "budget_capex_initial": "number|null",
  "budget_capex_landing": "number|null",
  "budget_capex_used": "number|null",
  "budget_opex": "number|null",
  "budget_opex_initial": "number|null",
  "budget_opex_landing": "number|null",
  "budget_opex_used": "number|null",
  "created_at": "datetime",
  "custom_attributes": { "attr_id": "value_id" },
  "description": "array (rich text)",
  "description_text": "string",
  "effort": "number|null",
  "effort_used": "number|null",
  "end_date": "date",
  "gain": "number|null",
  "gain_text": "string",
  "goals": "array",
  "importance": "string",
  "milestone_progress": "number (0-100)",
  "mood": "string (code)",
  "program": { "id", "name", "short_id", ... },
  "progress": "number (0-100)",
  "requesting_team": { "id", "name" },
  "risk": "string (code)",
  "start_date": "date",
  "teams": "array"
}
```

## Project Members Response

From `GET /projects/{id}/members/`:

```json
{
  "results": [
    {
      "user": { "id", "name", "initials", "picture", "current_position" },
      "role": { "id", "name", "description", "icon" }
    }
  ]
}
```

## Project Efforts Response

From `GET /projects/{id}/efforts/`:

```json
{
  "results": [
    {
      "team": { "id", "name" },
      "effort": "number|null",
      "effort_used": "number|null"
    }
  ]
}
```

## API Constraints

### Pagination
- Default page_size: 10
- Maximum page_size: **20**
- Response: `{ count, next, previous, results }`

### Rate Limits
- 15 calls per second
- 500 calls per minute
- 100,000 calls per day
- On 429: check `Retry-After` header

---

*Last updated: January 2026*
*Source: https://developers.airsaas.io/reference/introduction*
