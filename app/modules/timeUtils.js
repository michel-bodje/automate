/**
 * Time-related utilities for scheduling system
 */

// Constants
export const FIRM_TIMEZONE = "America/Toronto";

export const LUNCH_START_HOUR = 13; // 1pm
export const LUNCH_END_HOUR = 14;   // 2pm

/**
 * Checks if a time slot overlaps with the lunch break (1pm-2pm)
 * @param {Date} slotStart - Start time of the slot
 * @param {Date} slotEnd - End time of the slot
 * @returns {boolean} - True if the slot overlaps with lunch
 */
export function overlapsLunch(slotStart, slotEnd) {
  const lunchStart = new Date(slotStart);
  lunchStart.setHours(LUNCH_START_HOUR, 0, 0, 0);
  
  const lunchEnd = new Date(slotStart);
  lunchEnd.setHours(LUNCH_END_HOUR, 0, 0, 0);
  
  return (
    (slotStart < lunchEnd && slotEnd > lunchStart) ||
    (slotStart.getHours() === LUNCH_START_HOUR && slotStart.getMinutes() > 0) ||
    (slotEnd.getHours() === LUNCH_END_HOUR && slotEnd.getMinutes() > 0)
  );
}

/**
 * Adjusts a time slot to avoid lunch break if needed
 * @param {Date} slotStart - Original start time
 * @param {Date} slotEnd - Original end time
 * @returns {{start: Date, end: Date}|null} - Adjusted slot or null if not possible
 */
export function adjustForLunch(slotStart, slotEnd, slotDuration) {
  if (!overlapsLunch(slotStart, slotEnd)) {
    return { start: slotStart, end: slotEnd };
  }
  
  // Try moving to after lunch (2pm)
  const afterLunchStart = new Date(slotStart);
  afterLunchStart.setHours(LUNCH_END_HOUR, 0, 0, 0);
  const afterLunchEnd = new Date(afterLunchStart.getTime() + slotDuration);
  
  return { start: afterLunchStart, end: afterLunchEnd };
}