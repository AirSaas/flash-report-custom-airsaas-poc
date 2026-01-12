# Fetch AirSaas Project Data

Fetch all project data from AirSaas API and save to cache.

## Instructions

1. **Load Configuration**
   - Read `.env` for API credentials (AIRSAAS_API_KEY, AIRSAAS_BASE_URL)
   - Read `config/projects.json` for workspace and project list
   - Verify credentials are configured

2. **Fetch Reference Data** (cache for label resolution)
   - `GET /projects_moods/` - Mood definitions
   - `GET /projects_statuses/` - Status definitions
   - `GET /projects_risks/` - Risk definitions

3. **For Each Project in projects.json**
   Fetch all related data:
   - `GET /projects/{id}/?expand=owner,program,goals,teams,requesting_team` - Main project info
   - `GET /milestones/?expand=project&project={id}` - Milestones (filter by project)
   - `GET /decisions/?expand=owner,decision_maker,project&project={id}` - Decisions
   - `GET /attention_points/?expand=owner,project&project={id}` - Attention points

4. **Handle Pagination**
   - Check for `next` field in responses
   - Follow pagination links to get all results
   - Use `page_size=20` (maximum allowed)

5. **Save Results**
   - Combine all data into single JSON structure
   - Save to `data/{YYYY-MM-DD}_projects.json`
   - Include fetch timestamp and project count

6. **Report Results**
   - Number of projects fetched
   - Any errors encountered
   - Path to saved file

## API Authentication

Header: `Authorization: Api-Key {AIRSAAS_API_KEY}`

## Output Format

```json
{
  "fetched_at": "2025-01-15T10:30:00Z",
  "workspace": "aqme-corp-",
  "reference_data": {
    "moods": [...],
    "statuses": [...],
    "risks": [...]
  },
  "projects": [
    {
      "id": "uuid",
      "project": {...},
      "milestones": [...],
      "decisions": [...],
      "attention_points": [...]
    }
  ]
}
```

## Rate Limits

- 15 calls per second
- 500 calls per minute
- 100,000 calls per day
- On 429 error, wait for `Retry-After` header seconds

## Error Handling

- Log errors to `tracking/CLAUDE_ERRORS.md`
- Continue fetching other projects if one fails
- Report which projects failed at the end
