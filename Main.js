/**
 * CONTACT CENTER TEAM LEAD DASHBOARD
 * Version 1.0
 */

/**
 * CONTACT CENTER DASHBOARD
 */

function INSTALL(){

  buildSettings();

  buildRoster();

  buildDailyOperations();

  buildDashboardSheet();

  applyDropdowns();

  applyAgentDropdown();

  applyRosterValidation();

  applyFormatting();

  applyFormulas();

  createNamedRanges();

  updateBreakMonitor();

  onOpen();

}