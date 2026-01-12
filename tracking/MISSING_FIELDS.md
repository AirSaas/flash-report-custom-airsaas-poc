# Missing Fields Log

This document tracks AirSaas API fields that are needed for the PPT template but are not available through the public API.

## Known Missing Fields

| Field | PPT Location | Reason | Workaround |
|-------|--------------|--------|------------|
| Mood comment | Project Card | API returns mood code only, no text | Leave empty or use mood label |
| Budget | Project Card | Budget endpoints not in public API docs | Leave empty or manual input |
| EAC budget | Project Card | Not available in API | Leave empty |
| Team efforts | Project Planning | Efforts endpoint not in public API docs | Leave empty |
| Completion % | Project Progress | No direct API field | Calculate from milestones (completed/total) |
| Deployment area | Project Card | Field not identified in API | Leave empty |
| End users (actual) | Project Progress | Field not identified in API | Leave empty |
| End users (target) | Project Progress | Field not identified in API | Leave empty |

## Available Endpoints (per API docs)

### Documented Endpoints
| Endpoint | Expandable Fields |
|----------|-------------------|
| `GET /projects/` | owner, program, goals, teams, requesting_team |
| `GET /projects/{id}/` | owner, program, goals, teams, requesting_team |
| `GET /milestones/` | owner, team, project |
| `GET /decisions/` | owner, decision_maker, project, workspace_program |
| `GET /attention_points/` | owner, project, workspace_program |
| `GET /projects_moods/` | - |
| `GET /projects_statuses/` | - |
| `GET /projects_risks/` | - |

### NOT in Public API Docs
These endpoints were assumed but are **not documented** in the public API:
- `/projects/{id}/members/`
- `/projects/{id}/efforts/`
- `/projects/{id}/budget_lines/`
- `/projects/{id}/budget_values/`

## API Constraints

### Pagination
- Default page_size: 10
- Maximum page_size: **20** (not 100)
- Response contains: `count`, `next`, `previous`, `results`

### Rate Limits
- 15 calls per second
- 500 calls per minute
- 100,000 calls per day
- On 429: check `Retry-After` header

## Field Discovery Process

When a new missing field is identified:
1. Add entry to this table
2. Document API endpoint investigated
3. Note any potential workaround
4. Update `config/mapping.json` with status "missing"

---

*Last updated: January 2025*
*Source: https://developers.airsaas.io/reference/introduction*
