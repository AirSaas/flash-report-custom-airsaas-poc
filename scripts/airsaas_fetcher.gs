/**
 * AirSaas Data Fetcher - Google Apps Script Reference
 *
 * This script fetches project information from the AirSaas API.
 * Use as reference for API structure or standalone in Google Sheets.
 *
 * Features:
 * - Fetches: project info, milestones, decisions, attention points
 * - Fetches: workspace references (moods, statuses, risks)
 * - Resolves codes to human-readable labels
 * - Exports to Google Sheets with formatted tabs
 *
 * API Documentation: https://developers.airsaas.io/reference/introduction
 *
 * Setup:
 * 1. Open Google Sheets
 * 2. Extensions > Apps Script
 * 3. Paste this code
 * 4. Set API_KEY in script properties or directly below
 * 5. Run fetchAllProjectData()
 *
 * Authentication: Authorization: Api-Key {YOUR_API_KEY}
 * Base URL: https://api.airsaas.io/v1
 *
 * Rate Limits:
 * - 15 calls per second
 * - 500 calls per minute
 * - 100,000 calls per day
 *
 * Pagination:
 * - Default page_size: 10
 * - Maximum page_size: 20
 */

// Configuration
const CONFIG = {
  API_KEY: '', // Set via Script Properties or paste here
  BASE_URL: 'https://api.airsaas.io/v1',
  WORKSPACE: 'aqme-corp-',
  PAGE_SIZE: 20 // Max allowed by API
};

/**
 * Main entry point - fetches reference data for the workspace
 * These are used to resolve codes (mood, status, risk) to labels
 */
function fetchReferenceData() {
  const apiKey = CONFIG.API_KEY || PropertiesService.getScriptProperties().getProperty('AIRSAAS_API_KEY');

  if (!apiKey) {
    throw new Error('API key not configured. Set AIRSAAS_API_KEY in Script Properties.');
  }

  // Fetch reference data (documented endpoints)
  const referenceData = {
    moods: fetchAllPages('/projects_moods/', apiKey),
    statuses: fetchAllPages('/projects_statuses/', apiKey),
    risks: fetchAllPages('/projects_risks/', apiKey)
  };

  Logger.log('Reference data loaded');
  Logger.log('Moods: ' + referenceData.moods.length);
  Logger.log('Statuses: ' + referenceData.statuses.length);
  Logger.log('Risks: ' + referenceData.risks.length);

  return referenceData;
}

/**
 * Fetch complete data for a single project
 *
 * Documented endpoints used:
 * - GET /projects/{id}/?expand=owner,program,goals,teams,requesting_team
 * - GET /milestones/?project={id}&expand=owner,team,project
 * - GET /decisions/?project={id}&expand=owner,decision_maker,project
 * - GET /attention_points/?project={id}&expand=owner,project
 *
 * @param {string} projectId - UUID of the project
 * @param {string} apiKey - API key for authentication
 * @param {Object} referenceData - Pre-fetched reference data for label resolution
 */
function fetchProjectData(projectId, apiKey, referenceData) {
  const data = {
    id: projectId,
    fetchedAt: new Date().toISOString()
  };

  // Main project info with all expandable fields
  data.project = fetchJson(
    `/projects/${projectId}/?expand=owner,program,goals,teams,requesting_team`,
    apiKey
  );

  // Related data - using documented endpoints with project filter
  data.milestones = fetchAllPages(`/milestones/?project=${projectId}&expand=owner,team,project`, apiKey);
  data.decisions = fetchAllPages(`/decisions/?project=${projectId}&expand=owner,decision_maker,project`, apiKey);
  data.attentionPoints = fetchAllPages(`/attention_points/?project=${projectId}&expand=owner,project`, apiKey);

  // Resolve codes to labels
  if (referenceData) {
    data.resolved = {
      status: resolveLabel(data.project.status, referenceData.statuses),
      mood: resolveLabel(data.project.mood, referenceData.moods),
      risk: resolveLabel(data.project.risk, referenceData.risks)
    };
  }

  // Calculate completion percentage from milestones
  if (data.milestones && data.milestones.length > 0) {
    const completed = data.milestones.filter(m => m.is_completed).length;
    data.completionPercent = Math.round((completed / data.milestones.length) * 100);
  } else {
    data.completionPercent = null;
  }

  return data;
}

/**
 * Fetch JSON from API endpoint
 * @param {string} endpoint - API endpoint path
 * @param {string} apiKey - API key
 */
function fetchJson(endpoint, apiKey) {
  const url = CONFIG.BASE_URL + endpoint;

  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Api-Key ' + apiKey,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();

  if (code === 429) {
    // Rate limited - check Retry-After header
    const retryAfter = response.getHeaders()['Retry-After'] || 5;
    Logger.log('Rate limited. Retry after ' + retryAfter + ' seconds');
    Utilities.sleep(retryAfter * 1000);
    return fetchJson(endpoint, apiKey); // Retry
  }

  if (code !== 200) {
    Logger.log('Error fetching ' + url + ': ' + code);
    Logger.log(response.getContentText());
    throw new Error('API error: ' + code + ' - ' + response.getContentText());
  }

  return JSON.parse(response.getContentText());
}

/**
 * Fetch all pages from a paginated endpoint
 *
 * AirSaas pagination:
 * - Response: { count, next, previous, results }
 * - Max page_size: 20
 * - Follow 'next' URL until null
 *
 * @param {string} endpoint - API endpoint path
 * @param {string} apiKey - API key
 */
function fetchAllPages(endpoint, apiKey) {
  let results = [];
  let url = endpoint;

  // Add page_size if not already in URL
  if (!url.includes('page_size')) {
    url += (url.includes('?') ? '&' : '?') + 'page_size=' + CONFIG.PAGE_SIZE;
  }

  while (url) {
    const data = fetchJson(url, apiKey);

    if (Array.isArray(data)) {
      // Some endpoints return array directly (no pagination)
      results = results.concat(data);
      url = null;
    } else if (data.results) {
      // Paginated response
      results = results.concat(data.results);
      if (data.next) {
        // Extract path from full URL
        url = data.next.replace(CONFIG.BASE_URL, '');
      } else {
        url = null;
      }
    } else {
      // Single object response
      results.push(data);
      url = null;
    }
  }

  return results;
}

/**
 * Resolve a code to its human-readable label
 * @param {string} code - The code to resolve
 * @param {Array} referenceList - List of reference objects with code and label
 */
function resolveLabel(code, referenceList) {
  if (!code || !referenceList) return code;

  const match = referenceList.find(item => item.code === code || item.id === code);
  return match ? (match.label || match.name || code) : code;
}

/**
 * Export project data to a new sheet
 * @param {Object} projectData - Data from fetchProjectData
 */
function exportToSheet(projectData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projectName = projectData.project.name || projectData.id;

  // Create or get sheet
  let sheet = ss.getSheetByName(projectName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(projectName);
  }

  let row = 1;

  // Project Info
  sheet.getRange(row, 1).setValue('PROJECT INFO').setFontWeight('bold');
  row++;

  const projectFields = [
    ['Name', projectData.project.name],
    ['Short ID', projectData.project.short_id],
    ['Status', projectData.resolved?.status || projectData.project.status],
    ['Mood', projectData.resolved?.mood || projectData.project.mood],
    ['Risk', projectData.resolved?.risk || projectData.project.risk],
    ['Owner', projectData.project.owner?.full_name || ''],
    ['Program', projectData.project.program?.name || ''],
    ['Start Date', projectData.project.start_date],
    ['End Date', projectData.project.end_date],
    ['Completion %', projectData.completionPercent !== null ? projectData.completionPercent + '%' : 'N/A'],
    ['Description', projectData.project.description]
  ];

  projectFields.forEach(([field, value]) => {
    sheet.getRange(row, 1).setValue(field);
    sheet.getRange(row, 2).setValue(value || '');
    row++;
  });

  row++;

  // Goals (if expanded)
  if (projectData.project.goals && projectData.project.goals.length > 0) {
    sheet.getRange(row, 1).setValue('GOALS').setFontWeight('bold');
    row++;

    projectData.project.goals.forEach(goal => {
      sheet.getRange(row, 1).setValue(goal.name || goal.title || goal);
      row++;
    });
    row++;
  }

  // Milestones
  if (projectData.milestones && projectData.milestones.length > 0) {
    sheet.getRange(row, 1).setValue('MILESTONES').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 4).setValues([['Name', 'Date', 'Completed', 'Type']]);
    row++;

    projectData.milestones.forEach(m => {
      sheet.getRange(row, 1, 1, 4).setValues([[
        m.name || '',
        m.date || '',
        m.is_completed ? 'Yes' : 'No',
        m.milestone_type || ''
      ]]);
      row++;
    });
    row++;
  }

  // Decisions
  if (projectData.decisions && projectData.decisions.length > 0) {
    sheet.getRange(row, 1).setValue('DECISIONS').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 4).setValues([['Title', 'Status', 'Decision Maker', 'Date']]);
    row++;

    projectData.decisions.forEach(d => {
      sheet.getRange(row, 1, 1, 4).setValues([[
        d.title || '',
        d.status || '',
        d.decision_maker?.full_name || '',
        d.created_at || ''
      ]]);
      row++;
    });
    row++;
  }

  // Attention Points
  if (projectData.attentionPoints && projectData.attentionPoints.length > 0) {
    sheet.getRange(row, 1).setValue('ATTENTION POINTS').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 4).setValues([['Title', 'Severity', 'Status', 'Owner']]);
    row++;

    projectData.attentionPoints.forEach(ap => {
      sheet.getRange(row, 1, 1, 4).setValues([[
        ap.title || '',
        ap.severity || '',
        ap.status || '',
        ap.owner?.full_name || ''
      ]]);
      row++;
    });
    row++;
  }

  // Auto-resize columns
  sheet.autoResizeColumns(1, 4);

  Logger.log('Exported to sheet: ' + projectName);
}

/**
 * Fetch all projects in the workspace
 * Returns basic project list (use fetchProjectData for full details)
 */
function fetchAllProjects() {
  const apiKey = CONFIG.API_KEY || PropertiesService.getScriptProperties().getProperty('AIRSAAS_API_KEY');

  if (!apiKey) {
    throw new Error('API key not configured.');
  }

  const projects = fetchAllPages('/projects/?expand=owner,program', apiKey);
  Logger.log('Found ' + projects.length + ' projects');

  // Log project list
  projects.forEach(p => {
    Logger.log('- ' + p.name + ' (ID: ' + p.id + ')');
  });

  return projects;
}

/**
 * Test function - fetch a single project
 */
function testFetchProject() {
  const apiKey = CONFIG.API_KEY || PropertiesService.getScriptProperties().getProperty('AIRSAAS_API_KEY');
  const testProjectId = 'YOUR_PROJECT_UUID_HERE'; // Replace with actual ID

  const referenceData = fetchReferenceData();
  const projectData = fetchProjectData(testProjectId, apiKey, referenceData);

  Logger.log(JSON.stringify(projectData, null, 2));
  exportToSheet(projectData);
}

/**
 * Export multiple projects to sheets
 * @param {Array} projectIds - Array of project UUIDs
 */
function exportMultipleProjects(projectIds) {
  const apiKey = CONFIG.API_KEY || PropertiesService.getScriptProperties().getProperty('AIRSAAS_API_KEY');
  const referenceData = fetchReferenceData();

  projectIds.forEach(projectId => {
    try {
      const projectData = fetchProjectData(projectId, apiKey, referenceData);
      exportToSheet(projectData);
      Logger.log('Exported: ' + projectData.project.name);
    } catch (e) {
      Logger.log('Error exporting project ' + projectId + ': ' + e.message);
    }
  });
}

/**
 * Menu setup for Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AirSaas')
    .addItem('List All Projects', 'fetchAllProjects')
    .addItem('Fetch Reference Data', 'fetchReferenceData')
    .addItem('Fetch Test Project', 'testFetchProject')
    .addToUi();
}
