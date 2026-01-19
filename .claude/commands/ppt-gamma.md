# Generate PPT via Gamma API

Generate a PowerPoint presentation using the Gamma API. Claude Code builds the prompt dynamically based on `mapping.json` and fetched data.

## Prerequisites

1. **Fetched data**: Run `/fetch` first to get project data
2. **Gamma API key**: Set `GAMMA_API_KEY` in `.env` file
3. **Gamma Template ID**: Set `GAMMA_TEMPLATE_ID` in `.env` file

## Configuration

Environment variables in `.env`:

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `GAMMA_API_KEY` | Yes | - | Your Gamma API key |
| `GAMMA_BASE_URL` | No | `https://public-api.gamma.app/v1.0` | API base URL |
| `GAMMA_TEMPLATE_ID` | Yes | - | Template gamma ID from Gamma app |

## Output

| File | Description |
|------|-------------|
| `outputs/{date}_portfolio_gamma.pptx` | Generated PowerPoint presentation |
| `outputs/gamma_prompt.md` | Prompt sent to Gamma API (for debugging/audit) |
| `outputs/gamma_generation.json` | Initial API response with `generationId` (for resuming if interrupted) |
| `outputs/gamma_result.json` | Final API response with URLs, credits, and Gamma edit link |

**Online editing:** The `gammaUrl` in `gamma_result.json` links to the Gamma app for online editing.

---

## Instructions for Claude Code

When the user runs `/ppt-gamma`, follow these steps:

### Step 1: Load Configuration

1. Read `.env` for Gamma credentials:
   - `GAMMA_API_KEY` (required)
   - `GAMMA_BASE_URL` (default: https://public-api.gamma.app/v1.0)
   - `GAMMA_TEMPLATE_ID` (required)

2. Verify credentials are configured. If missing, prompt user to add them.

### Step 2: Load Data

1. Find the latest fetched data file in `data/` folder (format: `YYYY-MM-DD_projects.json`)
2. If no data found, prompt user to run `/fetch` first
3. Load `config/mapping.json` for field definitions

### Step 3: Build the Prompt Dynamically

Generate a markdown prompt based on the **PPT Output Structure** from SPEC.md section 5:

#### Slide Structure (per SPEC.md 5.1):

| Slide | Content |
|-------|---------|
| 1 | **Summary** - Table of all projects with ID, Name, Status, Mood, Owner |
| 2 to N | **Project slides** - One per project with all mapped fields |
| Last | **Data Notes** - List of unfilled fields and API limitations |

#### Prompt Template:

```markdown
# Portfolio Flash Report
**Date:** {YYYY-MM-DD}

---

## Summary

| ID | Project Name | Status | Mood | Owner |
|----|--------------|--------|------|-------|
{for each project: | {short_id} | {name} | {status} | {mood} | {owner} |}

---

{for each project:}

## {project.short_id}: {project.name}

### Overview
- **Status:** {resolved.status}
- **Mood:** {resolved.mood}
- **Risk:** {resolved.risk}
- **Owner:** {project.owner.name}
- **Program:** {project.program.name}

### Timeline
- **Start Date:** {project.start_date}
- **End Date:** {project.end_date}
- **Progress:** {project.progress}%

### Description
{project.description_text}

### Scope & Milestones
{for each milestone:}
- {milestone.name} ({milestone.date}) - {milestone.status}
{end for}

**Completed:** {count completed} / {count total}

### Budget
- **BAC (Initial):** {project.budget_capex_initial}€
- **Actual:** {project.budget_capex_used}€
- **EAC (Landing):** {project.budget_capex_landing}€

### Effort
- **Planned:** {project.effort} days
- **Used:** {project.effort_used} days

### Team Efforts
{for each effort in efforts:}
- {effort.team.name}: {effort.effort} planned / {effort.effort_used} used
{end for}

### Decisions
{for each decision in decisions:}
- {decision.title} ({decision.status})
{end for}

### Risks & Attention Points
{for each attention_point:}
- {attention_point.title} - {attention_point.severity}
{end for}

---

{end for each project}

## Data Notes

### API Limitations (fields not available)
- Mood comment (API returns code only, no free text)
- Deployment area (field not in API)
- End users actual/target (field not in API)

### Unfilled Fields per Project
{list any fields that were empty or null for each project}
```

### Step 4: Call Gamma API

1. **Create Generation Request:**
   ```bash
   curl -X POST "https://public-api.gamma.app/v1.0/generations/from-template" \
     -H "X-API-KEY: {GAMMA_API_KEY}" \
     -H "Content-Type: application/json" \
     -d '{
       "gammaId": "{GAMMA_TEMPLATE_ID}",
       "prompt": "{generated_prompt}",
       "exportAs": "pptx"
     }'
   ```

2. **Poll for Completion:**
   ```bash
   curl "https://public-api.gamma.app/v1.0/generations/{generationId}" \
     -H "X-API-KEY: {GAMMA_API_KEY}"
   ```

   Poll every 3 seconds until `status` is `completed` or `failed`.

3. **Download PPTX:**
   When completed, download from `exportUrl` and save to `outputs/{date}_portfolio_gamma.pptx`

### Step 5: Report Results

Report to the user:
- Output file path
- Number of projects included
- Gamma URL for online editing
- Credits used/remaining
- List of unfilled fields (if any)

---

## Gamma API Reference

### Primary Endpoint: Create from Template

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations/from-template`

**Headers:**
```
X-API-KEY: {GAMMA_API_KEY}
Content-Type: application/json
```

**Request (VERIFIED WORKING):**
```json
{
  "gammaId": "template-id-from-gamma",
  "prompt": "Your markdown content...",
  "exportAs": "pptx"
}
```

### CRITICAL: Parameters NOT Supported for from-template Endpoint

⚠️ **DO NOT USE these parameters with `/generations/from-template`:**

| Parameter | Error | Notes |
|-----------|-------|-------|
| `imageOptions` | `imageOptions.property source should not exist` | Not supported for template-based generation |
| `imageOptions.source` | Validation error 400 | Use only with `/generations` standard endpoint |
| `textMode` | Not applicable | Only for standard generation |
| `cardSplit` | Not applicable | Only for standard generation |
| `numCards` | Not applicable | Only for standard generation |

**The `/generations/from-template` endpoint only accepts:**
- `gammaId` (required) - Template ID
- `prompt` (required) - Instructions/content for the template
- `exportAs` (optional) - `"pptx"` or `"pdf"`
- `themeId` (optional) - Override theme styling

### Response (201 Created)

```json
{
  "generationId": "abc123"
}
```

### Poll for Completion

**Endpoint:** `GET https://public-api.gamma.app/v1.0/generations/{generationId}`

**Poll Settings:**
- Poll interval: 3 seconds
- Max attempts: 60 (3 minutes timeout)
- Status values: `pending`, `completed`, `failed`

**Completed Response:**
```json
{
  "generationId": "abc123",
  "status": "completed",
  "gammaUrl": "https://gamma.app/docs/xxx",
  "exportUrl": "https://assets.api.gamma.app/export/pptx/xxx/Untitled.pptx",
  "credits": {
    "deducted": 40,
    "remaining": 7735
  }
}
```

### Standard Generation Endpoint (Alternative)

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations`

This endpoint supports more options but does NOT use templates:

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

### Endpoints Summary

| Endpoint | Method | Use Case | Supports imageOptions |
|----------|--------|----------|----------------------|
| `/v1.0/generations/from-template` | POST | Generate from existing template | ❌ NO |
| `/v1.0/generations` | POST | Generate from text (standard) | ✅ YES |
| `/v1.0/generations/{id}` | GET | Poll generation status | N/A |
| `/v1.0/themes` | GET | List available themes | N/A |
| `/v1.0/folders` | GET | List folders | N/A |

### API Access

- Available on Pro, Ultra, Teams, Business plans
- Credit-based billing (~40 credits per generation)
- Full docs: https://developers.gamma.app/docs/getting-started

---

## Error Handling

- Log errors to `tracking/CLAUDE_ERRORS.md`
- If API key missing → prompt user to configure `.env`
- If no data → prompt user to run `/fetch`
- If Gamma API fails → show error message and suggest retry
- If polling times out (>5 minutes) → report timeout error
