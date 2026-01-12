# Interactive Field Mapping

Map AirSaas API fields to PPT template placeholders through interactive conversation.

## Instructions

### Phase 1: Template Analysis

1. **Check for Template File**
   - Look in `templates/` folder for PPT files
   - If no template exists, ask user to provide one or describe desired structure

2. **Analyze Template Structure**
   - If PPT file exists, describe the slide structure found
   - Identify placeholders/fields that need data
   - List the slides per project (e.g., Card, Progress, Planning)

3. **Present Template Summary**
   ```
   Template Analysis:
   - Summary Slide: [fields found]
   - Per-Project Slides:
     - Slide 1 (Card): [fields]
     - Slide 2 (Progress): [fields]
     - Slide 3 (Planning): [fields]
   - Data Notes Slide: automatic
   ```

### Phase 2: API Field Inventory

1. **Load Available API Fields**
   Read from cached data or document known fields:

   **Project Core:**
   - name, short_id, description
   - status, mood, risk (codes)
   - owner.full_name, program.name
   - start_date, end_date
   - goals[]

   **Related Data:**
   - members[]: user info, role
   - efforts[]: team, period, value
   - milestones[]: name, date, is_completed
   - decisions[]: title, status, date
   - attention_points[]: title, severity, status
   - budget_lines[]: name, type
   - budget_values[]: line, period, value

2. **Present API Fields**
   Show organized list of available fields

### Phase 3: Interactive Mapping

For each PPT placeholder:

1. **Propose a Match**
   ```
   PPT Field: "Project Status"
   Suggested API Source: project.status (with status label lookup)
   Is this correct? [Yes/No/Different field]
   ```

2. **Handle User Response**
   - **Yes**: Mark as "ok" in mapping
   - **No + alternative**: Use user's field, mark as "ok"
   - **Not available**: Mark as "missing", add to MISSING_FIELDS.md

3. **Ask About Transforms**
   If field needs transformation:
   ```
   This field needs formatting. Options:
   - format_date: Convert ISO date to readable format
   - resolve_label: Look up code in reference data
   - calculate: Derive from other fields
   - custom: Describe needed transformation
   ```

### Phase 4: Save Results

1. **Update mapping.json**
   - Write all field mappings with sources and transforms
   - Include status for each field

2. **Update MISSING_FIELDS.md**
   - Add any newly identified missing fields
   - Note workarounds discussed

3. **Summary Report**
   ```
   Mapping Complete:
   - Fields mapped: X
   - Fields missing: Y
   - Fields needing manual input: Z

   See config/mapping.json for full mapping
   See tracking/MISSING_FIELDS.md for missing fields
   ```

## Conversation Style

- Be concise but clear
- Group related fields together
- Offer reasonable defaults
- Explain why certain fields can't be mapped
- Ask for clarification when source is ambiguous

## Example Dialogue

```
Claude: I found your template has 3 slides per project. Let's map the fields.

For the Project Card slide, I see a "Budget" placeholder.
Available budget fields:
- budget_lines: Individual line items
- budget_values: Values by period

Should I sum all budget_values for total budget? [Yes/No/Explain more]

User: Yes, sum them

Claude: Got it. I'll create a transform to sum budget_values.
Next field: "Completion %" - I don't see a direct API field for this.
Options:
1. Calculate from milestones (completed/total)
2. Leave empty
3. You provide manually

Which approach?
```
