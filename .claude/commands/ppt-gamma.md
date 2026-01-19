# Generate PPT via Gamma API

Generate a PowerPoint presentation using the Gamma API with a pre-defined template.

## Quick Start

Simply run the Python script:

```bash
python3 scripts/generate_ppt_gamma.py
```

## What it does

1. Loads the latest fetched data from `data/{date}_projects.json`
2. Formats project data into markdown content
3. Calls Gamma API with the template ID from `.env`
4. Polls for completion
5. Downloads the PPTX to `outputs/{date}_portfolio_gamma.pptx`

## Prerequisites

1. **Fetched data**: Run `/fetch` first to get project data
2. **Gamma API key**: Set in `.env` file:
   ```
   GAMMA_API_KEY=sk-gamma-xxx
   GAMMA_TEMPLATE_ID=g_9d4wnyvr02om4zk
   ```

## Configuration

Environment variables in `.env`:

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `GAMMA_API_KEY` | Yes | - | Your Gamma API key |
| `GAMMA_BASE_URL` | No | `https://public-api.gamma.app/v1.0` | API base URL |
| `GAMMA_TEMPLATE_ID` | No | `g_9d4wnyvr02om4zk` | Template gamma ID |

## Output

- **File**: `outputs/{date}_portfolio_gamma.pptx`
- **Online**: Link to Gamma app for editing

## Instructions for Claude

When the user runs `/ppt-gamma`:

1. Execute the script:
   ```bash
   python3 scripts/generate_ppt_gamma.py
   ```

2. Report the results to the user:
   - Output file path
   - Number of projects
   - Gamma URL for online editing
   - Credits used/remaining
   - Any unfilled fields

---

## Gamma API Reference (for manual troubleshooting)

### Primary Endpoint: Create from Template

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations/from-template`

**Headers:**
```
X-API-KEY: {GAMMA_API_KEY}
Content-Type: application/json
```

**Request:**
```json
{
  "gammaId": "g_9d4wnyvr02om4zk",
  "prompt": "Your markdown content...",
  "exportAs": "pptx"
}
```

### Response (201 Created)

```json
{
  "generationId": "abc123"
}
```

### Poll for Completion

**Endpoint:** `GET https://public-api.gamma.app/v1.0/generations/{generationId}`

**Completed Response:**
```json
{
  "generationId": "abc123",
  "status": "completed",
  "gammaUrl": "https://gamma.app/docs/xxx",
  "exportUrl": "https://assets.api.gamma.app/export/pptx/xxx/Untitled.pptx",
  "credits": {
    "deducted": 32,
    "remaining": 7875
  }
}
```

### Other Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/v1.0/generations` | POST | Generate from text (standard) |
| `/v1.0/themes` | GET | List available themes |
| `/v1.0/folders` | GET | List folders |

### API Access

- Available on Pro, Ultra, Teams, Business plans
- Credit-based billing
- v0.2 deprecated Jan 16, 2026

**Full docs:** https://developers.gamma.app/docs/getting-started
