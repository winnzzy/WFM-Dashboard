/**
 * ==========================================================
 * BUSINESS RULES
 * ==========================================================
 */

const BREAK_WINDOWS = {

  Morning: {
    start: 12,
    end: 14
  },

  Afternoon: {
    start: 15,
    end: 17
  },

  Night: {
    start: 3,
    end: 6
  }

};

function isWithinBreakWindow(shift){

  const rule = BREAK_WINDOWS[shift];

  if(!rule) return false;

  const now = new Date();

  const hour = now.getHours();

  return hour >= rule.start &&
         hour < rule.end;

}