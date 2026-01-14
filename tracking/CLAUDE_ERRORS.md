# Claude Code Errors Log

This document tracks errors encountered during development and usage to prevent repeating mistakes.

## Error Log Format

Each entry should include:
- **Date**: When the error occurred
- **Command**: What command/action triggered it
- **Error**: The error message
- **Solution**: How it was fixed
- **Prevention**: How to avoid in future

---

## Logged Errors

### 2026-01-13: Budget fields showing "N/A (API unavailable)"

- **Date**: 2026-01-13
- **Command**: `/ppt-skill`
- **Error**: All budget fields displayed "N/A (API unavailable)" even when data existed in API
- **Root Cause**: `scripts/generate_ppt.py` had hardcoded wrong field names:
  - Script used: `budget_capex`, `budget_opex`
  - Actual API fields: `budget_capex_initial`, `budget_capex_used`, `budget_capex_landing`
  - Script checked `is_completed` for milestones, but data uses `status == 'done'`
- **Solution**: Updated `generate_ppt.py` to use correct field names matching API data:
  - `budget_capex_initial` → BAC
  - `budget_capex_used` → Actual
  - `budget_capex_landing` → EAC
  - Changed milestone check from `is_completed` to `status == 'done'`
- **Prevention**:
  1. Always read actual data from `data/*_projects.json` before writing scripts
  2. The `/mapping` command must update BOTH `mapping.json` AND `generate_ppt.py`
  3. Added CRITICAL RULES to `/mapping` command to enforce this

### 2026-01-13: /mapping command not being interactive

- **Date**: 2026-01-13
- **Command**: `/mapping`
- **Error**: Command displayed mapping status instead of running interactively
- **Root Cause**: Command instructions weren't explicit enough about REQUIRING interaction
- **Solution**: Updated `.claude/commands/mapping.md` with:
  - Bold IMPORTANT note at top requiring interactive behavior
  - CRITICAL RULES section with 5 explicit rules
  - Requirement to use AskUserQuestion tool for each decision
- **Prevention**: Command now clearly states "MUST be interactive" and lists exact steps

### 2026-01-13: Template images/borders not preserved

- **Date**: 2026-01-13
- **Command**: `/ppt-skill`
- **Error**: Generated PPT lost images, borders, and colors from template
- **Root Cause**: Adding new slides from layout doesn't copy shapes from template slides
- **Solution**: Implemented `duplicate_slide()` function using XML manipulation:
  - Copy shapes via `deepcopy()` of spTree elements
  - Copy image relationships
  - Find shapes by position instead of name (names change after duplication)
- **Prevention**: Always use `duplicate_slide()` when cloning template slides

### 2026-01-13: Placeholder "Double click to edit" showing

- **Date**: 2026-01-13
- **Command**: `/ppt-skill`
- **Error**: Summary and Data Notes slides showed "Double click to edit" overlapping titles
- **Root Cause**: `set_shape_text(shape, "")` clears text but placeholder still shows default text
- **Solution**: Created `delete_placeholder_shapes()` to remove placeholders entirely:
  ```python
  sp = shape._element
  sp.getparent().remove(sp)
  ```
- **Prevention**: On Summary/Data Notes slides, always delete placeholders, don't just clear text

### 2026-01-13: Text overlapping section titles in template

- **Date**: 2026-01-13
- **Command**: `/ppt-skill`
- **Error**: Content in "Description & Expected Benefits" overlapped with section title
- **Root Cause**: Section titles are visual elements in template background, not part of text shapes
- **Solution**: Add leading newline to push content below visual titles:
  ```python
  set_shape_text(achievements_shape, "\n" + goals_text, font_size=CONTENT_FONT)
  ```
- **Prevention**: For shapes at (0.4, 3.0) and (0.4, 4.4), always add leading newline

### 2026-01-13: Font sizes too large in project slides

- **Date**: 2026-01-13
- **Command**: `/ppt-skill`
- **Error**: Text in project review slides was too large, causing overflow
- **Root Cause**: Using template's default font sizes which were designed for less content
- **Solution**: Implemented explicit font sizes:
  - Title: 14pt
  - Content: 9pt
  - Date: 8pt
- **Prevention**: Font sizes documented in ppt-skill.md and mapping.md

---

## Common Issues Reference

### API Authentication
- **Symptom**: 401 Unauthorized
- **Check**: Is `AIRSAAS_API_KEY` set in `.env`?
- **Check**: Is the key format correct? (should be plain key, not "Api-Key xxx")
- **Header format**: `Authorization: Api-Key {key}`

### Pagination
- **Symptom**: Missing data, only first page returned
- **Check**: Are you following the `next` link in responses?
- **Note**: Default page_size is 20, max is 100

### Project Not Found
- **Symptom**: 404 on project endpoint
- **Check**: Is the project ID a valid UUID?
- **Check**: Does the API key have access to this workspace?

### Gamma API
- **Symptom**: Generation fails or times out
- **Check**: Is `GAMMA_API_KEY` set correctly?
- **Check**: Is the input text properly formatted with `\n---\n` slide breaks?
- **Note**: Poll `GET /generations/{id}` until status is "completed"

### 2026-01-13: Gamma API 401 Invalid Token

- **Date**: 2026-01-13
- **Command**: `/ppt-gamma`
- **Error**: `{"error":"invalid_token","error_description":"The access token is invalid or expired","statusCode":401}`
- **Root Cause**: The Gamma API key in `.env` is invalid or expired
- **Solution**: User needs to:
  1. Log in to https://gamma.app
  2. Go to API settings and generate a new API key
  3. Update `GAMMA_API_KEY` in `.env` with the new key
- **Prevention**: Gamma API keys may expire - check validity before running /ppt-gamma

### python-pptx
- **Symptom**: Import error
- **Check**: Is python-pptx installed? `pip install python-pptx`
- **Check**: Python version compatibility (3.7+)

---

*Last updated: January 2026*
