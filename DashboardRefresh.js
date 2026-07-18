/**
 * ==========================================================
 * DashboardRefresh.gs
 * ==========================================================
 */

function refreshDashboard() {

  try {
    updateBreakMonitor();
  } catch (e) {
    Logger.log("Break Monitor Error: " + e);
  }

  try {
    updateQueueShare();
  } catch (e) {
    Logger.log("Queue Share Error: " + e);
  }

  try {
    updateLateReturns();
  } catch (e) {
    Logger.log("Late Returns Error: " + e);
  }

}