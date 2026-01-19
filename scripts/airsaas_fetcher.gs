/**
 * AirSaas Data Fetcher - Google Apps Script Reference
 *
 * This script fetches project information from the AirSaas API.
 * Use as reference for API structure or standalone in Google Sheets.
 *
 * Features:
 * - Fetches: project, members, efforts, milestones, decisions, attention points
 * - Fetches: all workspace references (moods, statuses, risks, teams, users, programs)
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
 *
 * Available Endpoints:
 * - /projects/ - List all projects
 * - /projects/{id}/ - Single project (expand: owner, program, goals, teams, requesting_team)
 * - /projects/{id}/members/ - Project team members with roles
 * - /projects/{id}/efforts/ - Per-team effort breakdown
 * - /milestones/?project={id} - Project milestones (expand: owner, team, project)
 * - /decisions/?project={id} - Project decisions (expand: owner, decision_maker, project)
 * - /attention_points/?project={id} - Project attention points (expand: owner, project)
 * - /projects_moods/ - Mood definitions
 * - /projects_statuses/ - Status definitions
 * - /projects_risks/ - Risk definitions
 * - /teams/ - All workspace teams
 * - /users/ - Workspace members
 * - /programs/ - Programs list
 * - /project_custom_attributes/ - Custom attribute definitions
 */

// ============================================================================
// CONFIGURATION - Edit these values or use Script Properties
// ============================================================================
const CONFIG = {
  // Authentication
  API_KEY: '',                              // AirSaas API key (or set AIRSAAS_API_KEY in Script Properties)

  // API Settings
  BASE_URL: 'https://api.airsaas.io/v1',    // AirSaas API base URL
  PAGE_SIZE: 20,                            // Max items per page (API max: 20)

  // Workspace
  WORKSPACE: 'aqme-corp-',                  // Your workspace slug

  // Rate Limiting
  RATE_LIMIT_RETRY_DEFAULT: 5,              // Default retry delay in seconds if Retry-After header missing

  // Endpoints - Reference Data
  ENDPOINTS: {
    PROJECTS: '/projects/',                 // List/get projects
    MILESTONES: '/milestones/',             // Project milestones
    DECISIONS: '/decisions/',               // Project decisions
    ATTENTION_POINTS: '/attention_points/', // Project attention points/risks
    MOODS: '/projects_moods/',              // Mood definitions (blocked, complicated, issues, good)
    STATUSES: '/projects_statuses/',        // Status definitions (idea, ongoing, paused, finished, canceled)
    RISKS: '/projects_risks/',              // Risk definitions (low, medium, high)
    TEAMS: '/teams/',                       // Workspace teams
    USERS: '/users/',                       // Workspace members
    PROGRAMS: '/programs/',                 // Programs list
    CUSTOM_ATTRIBUTES: '/project_custom_attributes/' // Custom attribute definitions
  },

  // Expandable Fields (use with ?expand=field1,field2)
  EXPAND: {
    PROJECT: 'owner,program,goals,teams,requesting_team',
    MILESTONE: 'owner,team,project',
    DECISION: 'owner,decision_maker,project',
    ATTENTION_POINT: 'owner,project'
  }
};

/**
 * Main entry point - fetches reference data for the workspace
 * These are used to resolve codes (mood, status, risk) to labels
 *
 * Documented endpoints:
 * - GET /projects_moods/ - Mood definitions (blocked, complicated, issues, good)
 * - GET /projects_statuses/ - Status definitions (idea, ongoing, paused, finished, canceled)
 * - GET /projects_risks/ - Risk definitions (low, medium, high)
 * - GET /teams/ - All workspace teams
 * - GET /users/ - Workspace members
 * - GET /programs/ - Programs list
 * - GET /project_custom_attributes/ - Custom attribute definitions
 */
function fetchReferenceData() {
  const apiKey = CONFIG.API_KEY || PropertiesService.getScriptProperties().getProperty('AIRSAAS_API_KEY');

  if (!apiKey) {
    throw new Error('API key not configured. Set AIRSAAS_API_KEY in Script Properties.');
  }

  // Fetch reference data (documented endpoints)
  const referenceData = {
    moods: fetchAllPages(CONFIG.ENDPOINTS.MOODS, apiKey),
    statuses: fetchAllPages(CONFIG.ENDPOINTS.STATUSES, apiKey),
    risks: fetchAllPages(CONFIG.ENDPOINTS.RISKS, apiKey),
    teams: fetchAllPages(CONFIG.ENDPOINTS.TEAMS, apiKey),
    users: fetchAllPages(CONFIG.ENDPOINTS.USERS, apiKey),
    programs: fetchAllPages(CONFIG.ENDPOINTS.PROGRAMS, apiKey),
    customAttributes: fetchAllPages(CONFIG.ENDPOINTS.CUSTOM_ATTRIBUTES, apiKey)
  };

  Logger.log('Reference data loaded');
  Logger.log('Moods: ' + referenceData.moods.length);
  Logger.log('Statuses: ' + referenceData.statuses.length);
  Logger.log('Risks: ' + referenceData.risks.length);
  Logger.log('Teams: ' + referenceData.teams.length);
  Logger.log('Users: ' + referenceData.users.length);
  Logger.log('Programs: ' + referenceData.programs.length);
  Logger.log('Custom Attributes: ' + referenceData.customAttributes.length);

  return referenceData;
}

/**
 * Fetch complete data for a single project
 *
 * Documented endpoints used:
 * - GET /projects/{id}/?expand=owner,program,goals,teams,requesting_team
 * - GET /projects/{id}/members/ - Project team members with roles
 * - GET /projects/{id}/efforts/ - Per-team effort breakdown
 * - GET /milestones/?project={id}&expand=owner,team,project
 * - GET /decisions/?project={id}&expand=owner,decision_maker,project
 * - GET /attention_points/?project={id}&expand=owner,project
 *
 * Undocumented endpoints (may not be available):
 * - GET /projects/{id}/budget_lines/ - Budget line items
 * - GET /projects/{id}/budget_values/ - Budget values per line
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
    `${CONFIG.ENDPOINTS.PROJECTS}${projectId}/?expand=${CONFIG.EXPAND.PROJECT}`,
    apiKey
  );

  // Project sub-resources
  data.members = fetchAllPages(`${CONFIG.ENDPOINTS.PROJECTS}${projectId}/members/`, apiKey);
  data.efforts = fetchAllPages(`${CONFIG.ENDPOINTS.PROJECTS}${projectId}/efforts/`, apiKey);

  // Budget data (undocumented endpoints - may not be available)
  try {
    data.budgetLines = fetchAllPages(`${CONFIG.ENDPOINTS.PROJECTS}${projectId}/budget_lines/`, apiKey);
    data.budgetValues = fetchAllPages(`${CONFIG.ENDPOINTS.PROJECTS}${projectId}/budget_values/`, apiKey);
  } catch (e) {
    // Budget endpoints are undocumented and may not exist for all workspaces
    Logger.log('Budget endpoints not available (undocumented API): ' + e.message);
    data.budgetLines = [];
    data.budgetValues = [];
  }

  // Related data - using documented endpoints with project filter
  data.milestones = fetchAllPages(`${CONFIG.ENDPOINTS.MILESTONES}?project=${projectId}&expand=${CONFIG.EXPAND.MILESTONE}`, apiKey);
  data.decisions = fetchAllPages(`${CONFIG.ENDPOINTS.DECISIONS}?project=${projectId}&expand=${CONFIG.EXPAND.DECISION}`, apiKey);
  data.attentionPoints = fetchAllPages(`${CONFIG.ENDPOINTS.ATTENTION_POINTS}?project=${projectId}&expand=${CONFIG.EXPAND.ATTENTION_POINT}`, apiKey);

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
    const retryAfter = response.getHeaders()['Retry-After'] || CONFIG.RATE_LIMIT_RETRY_DEFAULT;
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
    ['Owner', projectData.project.owner?.full_name || projectData.project.owner?.name || ''],
    ['Program', projectData.project.program?.name || ''],
    ['Start Date', projectData.project.start_date],
    ['End Date', projectData.project.end_date],
    ['Progress', projectData.project.progress !== null ? projectData.project.progress + '%' : 'N/A'],
    ['Milestone Progress', projectData.project.milestone_progress !== null ? projectData.project.milestone_progress + '%' : 'N/A'],
    ['Completion %', projectData.completionPercent !== null ? projectData.completionPercent + '%' : 'N/A'],
    ['Budget CAPEX (BAC)', projectData.project.budget_capex_initial],
    ['Budget CAPEX (Actual)', projectData.project.budget_capex_used],
    ['Budget CAPEX (EAC)', projectData.project.budget_capex_landing],
    ['Budget OPEX (BAC)', projectData.project.budget_opex_initial],
    ['Budget OPEX (Actual)', projectData.project.budget_opex_used],
    ['Budget OPEX (EAC)', projectData.project.budget_opex_landing],
    ['Effort (Planned)', projectData.project.effort],
    ['Effort (Used)', projectData.project.effort_used],
    ['Gain', projectData.project.gain],
    ['Description', projectData.project.description_text || projectData.project.description]
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
        d.decision_maker?.full_name || d.decision_maker?.name || '',
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
        ap.owner?.full_name || ap.owner?.name || ''
      ]]);
      row++;
    });
    row++;
  }

  // Team Members
  if (projectData.members && projectData.members.length > 0) {
    sheet.getRange(row, 1).setValue('TEAM MEMBERS').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 3).setValues([['Name', 'Role', 'Position']]);
    row++;

    projectData.members.forEach(m => {
      sheet.getRange(row, 1, 1, 3).setValues([[
        m.user?.full_name || m.user?.name || '',
        m.role?.name || '',
        m.user?.current_position || ''
      ]]);
      row++;
    });
    row++;
  }

  // Team Efforts
  if (projectData.efforts && projectData.efforts.length > 0) {
    sheet.getRange(row, 1).setValue('TEAM EFFORTS').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 3).setValues([['Team', 'Planned', 'Used']]);
    row++;

    projectData.efforts.forEach(e => {
      sheet.getRange(row, 1, 1, 3).setValues([[
        e.team?.name || '',
        e.effort || 0,
        e.effort_used || 0
      ]]);
      row++;
    });
    row++;
  }

  // Budget Lines
  if (projectData.budgetLines && projectData.budgetLines.length > 0) {
    sheet.getRange(row, 1).setValue('BUDGET LINES').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 4).setValues([['Name', 'Type', 'Amount', 'Currency']]);
    row++;

    projectData.budgetLines.forEach(bl => {
      sheet.getRange(row, 1, 1, 4).setValues([[
        bl.name || '',
        bl.type || '',
        bl.amount || 0,
        bl.currency || ''
      ]]);
      row++;
    });
    row++;
  }

  // Budget Values
  if (projectData.budgetValues && projectData.budgetValues.length > 0) {
    sheet.getRange(row, 1).setValue('BUDGET VALUES').setFontWeight('bold');
    row++;
    sheet.getRange(row, 1, 1, 4).setValues([['Line', 'Period', 'Value', 'Type']]);
    row++;

    projectData.budgetValues.forEach(bv => {
      sheet.getRange(row, 1, 1, 4).setValues([[
        bv.budget_line?.name || bv.budget_line || '',
        bv.period || '',
        bv.value || 0,
        bv.value_type || ''
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

  const projects = fetchAllPages(`${CONFIG.ENDPOINTS.PROJECTS}?expand=owner,program`, apiKey);
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
