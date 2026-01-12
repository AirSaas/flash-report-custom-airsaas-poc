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

*No errors logged yet.*

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

### python-pptx
- **Symptom**: Import error
- **Check**: Is python-pptx installed? `pip install python-pptx`
- **Check**: Python version compatibility (3.7+)

---

*Last updated: January 2025*
