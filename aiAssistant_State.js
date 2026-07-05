/* =========================
   SIDEBAR STATE / SETTINGS
========================= */

function getSidebarState() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = spreadsheet.getSheets().map(s => s.getName());

  const userProps = PropertiesService.getUserProperties();
  let selectedSheets = [];
  let sessionHistory = [];

  try {
    selectedSheets = JSON.parse(userProps.getProperty('AI_SELECTED_SHEETS') || '[]');
  } catch (e) {
    selectedSheets = [];
  }

  try {
    sessionHistory = JSON.parse(userProps.getProperty('AI_SESSION_HISTORY') || '[]');
  } catch (e) {
    sessionHistory = [];
  }

  selectedSheets = selectedSheets.filter(name => allSheetNames.includes(name));

  return {
    allSheetNames: allSheetNames,
    selectedSheets: selectedSheets,
    sessionHistory: sessionHistory
  };
}

function saveSidebarState(selectedSheets, sessionHistory) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const realSheetNames = spreadsheet.getSheets().map(s => s.getName());

  const safeSelectedSheets = Array.isArray(selectedSheets)
    ? selectedSheets.filter(name => realSheetNames.includes(name))
    : [];

  const safeSessionHistory = Array.isArray(sessionHistory) ? sessionHistory : [];

  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('AI_SELECTED_SHEETS', JSON.stringify(safeSelectedSheets));
  userProps.setProperty('AI_SESSION_HISTORY', JSON.stringify(safeSessionHistory));
  return true;
}

function clearSidebarState() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty('AI_SELECTED_SHEETS');
  userProps.deleteProperty('AI_SESSION_HISTORY');
  return true;
}

function getSheetNames() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheets().map(sheet => sheet.getName());
}

function updateSelectedSheets(selectedSheets) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const realSheetNames = spreadsheet.getSheets().map(s => s.getName());

  const safeSelectedSheets = Array.isArray(selectedSheets)
    ? selectedSheets.filter(name => realSheetNames.includes(name))
    : [];

  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('AI_SELECTED_SHEETS', JSON.stringify(safeSelectedSheets));
  return safeSelectedSheets;
}
