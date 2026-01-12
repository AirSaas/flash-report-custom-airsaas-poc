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

### Phase 3: Call Gamma API

1. **Prepare API Request**
   ```
   POST {GAMMA_BASE_URL}/generations
   Headers:
     X-API-KEY: {GAMMA_API_KEY}
     Content-Type: application/json

   Body:
   {
     "input": "{full_markdown_content}",
     "exportOptions": {
       "pptx": true
     }
   }
   ```

2. **Poll for Completion**
   ```
   GET {GAMMA_BASE_URL}/generations/{id}
   Headers:
     X-API-KEY: {GAMMA_API_KEY}
   ```
   - Poll every 5 seconds
   - Check for status "completed"
   - Get download URL from response

3. **Download PPTX**
   - Fetch the PPTX from download URL
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

## Gamma API Notes

- Slide breaks: Use `\n---\n` in markdown
- Formatting: Standard markdown supported
- Tables: Markdown tables work
- Images: Not supported in this flow
- Rate limits: Be aware of API quotas

## Content Formatting Tips

- Keep slides concise
- Use bullet points for lists
- Tables for structured data
- Bold for emphasis on key metrics
- Consistent heading hierarchy
