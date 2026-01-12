# Fetch Data and Generate All PPT Versions

Complete workflow: fetch fresh data and generate presentations with both engines.

## Instructions

This command combines `/fetch`, `/ppt-gamma`, and `/ppt-skill` into a single workflow.

### Step 1: Fetch Fresh Data

Execute the full `/fetch` workflow:
1. Load configuration from `.env` and `config/projects.json`
2. Fetch reference data (moods, statuses, risks)
3. Fetch all project data from AirSaas API
4. Save to `data/{date}_projects.json`

### Step 2: Generate Gamma PPT

Execute the `/ppt-gamma` workflow:
1. Load freshly fetched data
2. Apply field mappings
3. Generate markdown content
4. Call Gamma API
5. Download and save to `outputs/{date}_portfolio_gamma.pptx`

### Step 3: Generate python-pptx PPT

Execute the `/ppt-skill` workflow:
1. Load freshly fetched data
2. Apply field mappings
3. Generate presentation programmatically
4. Save to `outputs/{date}_portfolio_skill.pptx`

### Step 4: Comparison Report

Generate a summary comparing both outputs:

```
========================================
PPT Generation Complete
========================================

Data Fetched:
- Projects: {count}
- Timestamp: {datetime}
- File: data/{date}_projects.json

Gamma PPT:
- Output: outputs/{date}_portfolio_gamma.pptx
- Slides: {count}
- Status: Success/Failed

python-pptx PPT:
- Output: outputs/{date}_portfolio_skill.pptx
- Slides: {count}
- Status: Success/Failed

Unfilled Fields (both versions):
- {field1} in {project1}
- {field2} in {project2}

Comparison Notes:
- Gamma: AI-styled design, may vary from template
- python-pptx: Programmatic, matches template structure exactly
========================================
```

## Error Handling

If any step fails:
1. Log error to `tracking/CLAUDE_ERRORS.md`
2. Continue with remaining steps if possible
3. Report partial success in final summary

Example:
```
Gamma PPT generation failed: API rate limit exceeded
python-pptx generation succeeded
See tracking/CLAUDE_ERRORS.md for details
```

## Usage Notes

- This command takes longer as it runs the full pipeline
- Use when you need fresh data and both PPT versions
- For just fetching, use `/fetch`
- For just generating with existing data, use `/ppt-gamma` or `/ppt-skill`
