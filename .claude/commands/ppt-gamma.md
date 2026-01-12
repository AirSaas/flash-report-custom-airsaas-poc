# Generate PPT via Gamma API

Generate a PowerPoint presentation using the Gamma API (AI-powered design).

## Instructions

### Phase 1: Load Data

1. **Find Latest Fetched Data**
   - Look in `data/` folder for most recent `{date}_projects.json`
   - If no data found, prompt user to run `/fetch` first

2. **Load Mapping Configuration**
   - Read `config/mapping.json`
   - Identify fields and their sources/transforms

3. **Load Reference Data**
   - Extract moods, statuses, risks from fetched data
   - Build lookup tables for label resolution

### Phase 2: Prepare Content

1. **Build Summary Slide Content**
   ```markdown
   # Portfolio Flash Report
   **Date:** {current_date}

   | Project | Status | Mood |
   |---------|--------|------|
   | {name} | {status_label} | {mood_label} |
   ...
   ```

2. **Build Per-Project Slides**
   For each project, create slides based on mapping.json structure:

   ```markdown
   ---
   # {project_name}

   **Status:** {status} | **Mood:** {mood} | **Risk:** {risk}

   **Owner:** {owner}
   **Program:** {program}

   ## Key Information
   - Budget: {budget}
   - Timeline: {start_date} - {end_date}

   ## Achievements
   {goals_formatted}

   ---
   # {project_name} - Progress

   **Completion:** {completion_percent}

   ## Milestones
   {milestones_formatted}

   ---
   # {project_name} - Planning

   ## Team Efforts
   {efforts_table}

   ## Decisions
   {decisions_list}

   ## Attention Points
   {attention_points_list}
   ```

3. **Build Data Notes Slide**
   ```markdown
   ---
   # Data Notes

   The following fields could not be populated:
   {list_of_unfilled_fields}
   ```

4. **Track Unfilled Fields**
   - For each field with status "missing" or empty data
   - Add to unfilled list with project context

### Phase 3: Call Gamma API (v1.0)

1. **Prepare API Request**
   ```
   POST https://public-api.gamma.app/v1.0/generations
   Headers:
     Authorization: Bearer {GAMMA_API_KEY}
     Content-Type: application/json

   Body:
   {
     "inputText": "{full_markdown_content}",
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

2. **Key Parameters Explained**
   | Parameter | Value | Why |
   |-----------|-------|-----|
   | `textMode` | `preserve` | Keep exact text from inputText |
   | `format` | `presentation` | Generate slides |
   | `cardSplit` | `inputTextBreaks` | Use `---` as slide breaks |
   | `exportAs` | `pptx` | Get downloadable PPTX |
   | `cardOptions.dimensions` | `16x9` | Standard presentation ratio |
   | `imageOptions.source` | `noImages` | No AI images (use template style) |

3. **Response Handling**
   - Response includes `url` to the generated gamma
   - If `exportAs: "pptx"`, response includes `pptxUrl` for download
   - Download URLs expire after some time - save immediately

4. **Download PPTX**
   - Fetch the PPTX from `pptxUrl` in response
   - Save to `outputs/{date}_portfolio_gamma.pptx`

### Phase 4: Report Results

1. **Success Output**
   ```
   PPT Generated Successfully!

   Output: outputs/{date}_portfolio_gamma.pptx
   Projects: {count}
   Slides: {total_slides}

   Unfilled Fields:
   - {field1} in {project1}
   - {field2} in {project2}
   ```

2. **Error Handling**
   - Log errors to `tracking/CLAUDE_ERRORS.md`
   - Report what went wrong
   - Suggest fixes

## Gamma API v1.0 Reference

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations`

**Authentication:** `Authorization: Bearer {API_KEY}`

### Required Parameters

| Parameter | Type | Description |
|-----------|------|-------------|
| `inputText` | string | Content to generate (max 100,000 tokens) |
| `textMode` | string | `generate`, `condense`, or `preserve` |

### Optional Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `format` | string | `presentation` | `presentation`, `document`, `social`, `webpage` |
| `numCards` | integer | 10 | Number of slides (1-60 Pro, 1-75 Ultra) |
| `cardSplit` | string | `auto` | `auto` or `inputTextBreaks` (use `---`) |
| `exportAs` | string | - | `pdf` or `pptx` |
| `themeId` | string | workspace default | Theme ID from List Themes API |
| `additionalInstructions` | string | - | Extra instructions (1-2000 chars) |

### textOptions

| Parameter | Type | Default | Values |
|-----------|------|---------|--------|
| `amount` | string | `medium` | `brief`, `medium`, `detailed`, `extensive` |
| `tone` | string | - | Custom tone (1-500 chars) |
| `audience` | string | - | Target audience (1-500 chars) |
| `language` | string | `en` | ISO language code (`fr`, `es`, etc.) |

### cardOptions

| Parameter | Type | Default | Values |
|-----------|------|---------|--------|
| `dimensions` | string | `fluid` | Presentations: `fluid`, `16x9`, `4x3`. Documents: `pageless`, `letter`, `a4`. Social: `1x1`, `4x5`, `9x16` |
| `headerFooter` | object | - | Header/footer config with positions: `topLeft`, `topRight`, `topCenter`, `bottomLeft`, `bottomRight`, `bottomCenter` |

### imageOptions

| Parameter | Type | Default | Values |
|-----------|------|---------|--------|
| `source` | string | `aiGenerated` | `aiGenerated`, `pictographic`, `unsplash`, `giphy`, `webAllImages`, `webFreeToUse`, `webFreeToUseCommercially`, `placeholder`, `noImages` |
| `model` | string | auto | AI model for generation (when source is `aiGenerated`) |
| `style` | string | - | Artistic style (1-500 chars) |

### sharingOptions

| Parameter | Type | Default | Values |
|-----------|------|---------|--------|
| `workspaceAccess` | string | workspace default | `noAccess`, `view`, `comment`, `edit`, `fullAccess` |
| `externalAccess` | string | workspace default | `noAccess`, `view`, `comment`, `edit` |
| `emailOptions.recipients` | array | - | Email addresses for direct sharing |
| `emailOptions.access` | string | - | `view`, `comment`, `edit`, `fullAccess` |

### Text Modes Explained

| Mode | Use Case |
|------|----------|
| `generate` | Brief input, let AI expand |
| `condense` | Large input, summarize it |
| `preserve` | Keep exact text as-is |

### Slide Breaks

When using `cardSplit: "inputTextBreaks"`, use `\n---\n` (three dashes on its own line) to separate slides.

### Input Text Notes

- Max length: 100,000 tokens (~400,000 characters)
- Supports image URLs inline in text
- Use `\n---\n` for card breaks when `cardSplit: "inputTextBreaks"`

## Content Formatting Tips

- Keep slides concise
- Use bullet points for lists
- Tables for structured data
- Bold for emphasis on key metrics
- Consistent heading hierarchy

## Other Gamma APIs

### Create from Template API (Beta)

**Endpoint:** `POST https://public-api.gamma.app/v1.0/generations/from-template`

Create a new gamma based on an existing template.

| Parameter | Type | Required | Description |
|-----------|------|----------|-------------|
| `gammaId` | string | Yes | Template gamma ID to modify |
| `prompt` | string | Yes | Instructions for content modification (max 100k tokens) |
| `themeId` | string | No | Visual styling override |
| `folderIds` | array | No | Storage folders |
| `exportAs` | string | No | `pdf` or `pptx` |
| `imageOptions` | object | No | `model`, `style` |
| `sharingOptions` | object | No | Same as Generate API |

### List Themes API

**Endpoint:** `GET https://public-api.gamma.app/v1.0/themes`

| Parameter | Type | Description |
|-----------|------|-------------|
| `query` | string | Search by name (case-insensitive) |
| `limit` | integer | Items per page (max 50) |
| `after` | string | Cursor for pagination |

**Response:** `{ data: [{id, name, type, colorKeywords, toneKeywords}], hasMore, nextCursor }`

### List Folders API

**Endpoint:** `GET https://public-api.gamma.app/v1.0/folders`

Same parameters as List Themes. Returns `{ data: [{id, name}], hasMore, nextCursor }`

## API Access & Pricing

- Available on Pro, Ultra, Teams, Business plans
- Credit-based billing system
- v0.2 deprecated Jan 16, 2026 - use v1.0
- Create from Template API is in beta

**Docs:** https://developers.gamma.app/docs/getting-started
