import { 
  Lawyer,
  getLawyer,
  LUNCH_SLOT,
  RANGE_IN_DAYS,
} from "../index.js";

export const locationRules = {
  // Centralized list of locations
  locations: ["office", "phone", "teams"],
  
  // List of unavailability for each lawyer
  lawyerUnavailability: {
    DH: {
      office: ["Monday"],
    },
    TG: {
      office: ["Friday"],
    }
  },
};

/**
 * Creates a time slot given start and end times.
 * @param {Date} start - The start time of the event.
 * @param {Date} end - The end time of the event.
 * @param {string} location - The location of the event.
 * @returns {{ start: Date, end: Date, location: string }} - A time slot object: with start, end, and location properties.
 */
function createSlot(start, end, location) {
  return {
    start: start,
    end: end,
    location: location
  };
}

/**
 * Maps a Microsoft Graph event to a slot object.
 * @param {microsoftgraph.Event} event - The event to map.
 * @returns {{ start: Date, end: Date, location: string }} A slot object with start, end, and location properties.
 */
function mapEventToSlot(event) {
  return {
    start: new Date(event.start.dateTime),
    end: new Date(event.end.dateTime),
    location: event.location?.displayName || "",
  };
}

/**
 * Checks if a proposed slot conflicts with the lawyer's unavailability.
 * @param {string} lawyerId - The ID of the lawyer.
 * @param {{ start: Date, end: Date, location: string }} proposedSlot - The proposed time slot
 * @returns {boolean} True if there is a conflict, false otherwise.
 */
function hasAvailabilityConflict(lawyerId, proposedSlot) {
  const proposedDay = proposedSlot.start.toLocaleString("en-US", { weekday: "long" });
  const proposedLocation = proposedSlot.location;

  const unavailability = locationRules.lawyerUnavailability?.[lawyerId] || {};

  return Object.entries(unavailability).some(([location, days]) =>
    location === proposedLocation && days.includes(proposedDay)
  );
}

/**
 * Checks if a proposed slot has a location conflict with existing events.
 * A location conflict occurs when two events are scheduled in the office
 * at the same time.
 *
 * @param {{ start: Date, end: Date, location: string }} proposedSlot - The proposed time slot
 * @param {Array<microsoftgraph.Event>} allEvents - Array of existing events to check for conflicts.
 * @returns {boolean} True if there is a location conflict, false otherwise.
 */
function hasOfficeConflict(proposedSlot, allEvents) {
  try {
    const proposedLocation = proposedSlot.location;

    if (proposedLocation !== "office") {
      return false;
    }

    for (const event of allEvents) {
      const scheduledLocation = event.location?.displayName?.toLowerCase();
      const eventSlot = mapEventToSlot(event);
      if (scheduledLocation === "office" && isOverlapping(proposedSlot, eventSlot)) {
        // log the conflict
        console.warn(`Office conflict detected:`, proposedSlot, event);
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error("Error checking for office location conflict:", error);
    return false;
  }
}

/**
 * Checks if a proposed slot has a virtual conflict with existing events.
 * A virtual conflict occurs when two virtual meetings are scheduled with
 * Dorin Holban or Tim Gagin at the same time.
 *
 * @param {string} lawyerId - The ID of the lawyer.
 * @param {{ start: Date, end: Date, location: string }} proposedSlot - The proposed time slot
 * @param {Array<microsoftgraph.Event>} allEvents - Array of existing events to check for conflicts.
 * @returns {boolean} True if there is a virtual conflict, false otherwise.
 */
function hasVirtualConflict(lawyerId, proposedSlot, allEvents) {
  try {
    // If the lawyer is not DH or TG, return false
    if (!["DH", "TG"].includes(lawyerId)) {
      return false;
    }

    const otherLawyerId = lawyerId === "DH" ? "TG" : "DH"; 
    const otherLawyer = getLawyer(otherLawyerId);

    const proposedIsVirtual = isVirtualMeeting(proposedSlot);

    // Check if any existing event is scheduled in a virtual meeting 
    // with the other lawyer at the same time
    for (const event of allEvents) {
      const eventSlot = mapEventToSlot(event);
      if (
        event.categories?.includes(otherLawyer) &&
        isVirtualMeeting(eventSlot) &&
        isOverlapping(proposedSlot, eventSlot) &&
        proposedIsVirtual
      ) {
        // log the conflict
        console.warn(`Virtual conflict detected for ${proposedSlot} with ${event.subject}`);
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error("Error checking for DH_TG conflict:", error);
    // on fail, don't block scheduling. You'll have to check manually
    return false;
  }
}

/**
 * Checks if a proposed slot conflicts with the required break time
 * for the given lawyer. A break conflict occurs when there is not enough time
 * between the proposed appointment and either the previous or next appointment
 * with the same lawyer.
 *
 * @param {string} lawyerId - The ID of the lawyer.
 * @param {{ start: Date, end: Date, location: string }} proposedSlot - The proposed time slot
 * @param {Array<microsoftgraph.Event>} allEvents - Array of existing events to check against.
 * @returns {boolean} True if there is a break conflict, false otherwise.
 */
function hasBreakConflict(lawyerId, proposedSlot, allEvents) {
  const lawyer = getLawyer(lawyerId);
  const requiredBreak = lawyer.breakMinutes * (60 * 1000);

  // Filter events for the specific lawyer and sort by start time (ascending)
  const lawyerEvents = allEvents
    .filter(event => event.categories?.includes(lawyerId))
    .sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime));

  let previousEvent = null;
  let nextEvent = null;

  // Find the immediate previous and next events
  for (const event of lawyerEvents) {
    const eventSlot = mapEventToSlot(event);

    if (eventSlot.end <= proposedSlot.start) {
      previousEvent = eventSlot; // Update previous event
    } else if (eventSlot.start >= proposedSlot.end) {
      nextEvent = eventSlot; // Update next event
      break; // Stop searching once the next event is found
    }
  }

  // Check for conflict with the previous event
  if (previousEvent) {
    const breakTime = proposedSlot.start.getTime() - previousEvent.end.getTime();
    if (breakTime < requiredBreak) {
      console.warn(`Break conflict detected with previous event:`, proposedSlot, previousEvent);
      return true;
    }
  }

  // Check for conflict with the next event
  if (nextEvent) {
    const breakTime = nextEvent.start.getTime() - proposedSlot.end.getTime();
    if (breakTime < requiredBreak) {
      console.warn(`Break conflict detected with next event:`, proposedSlot, nextEvent);
      return true;
    }
  }

  return false;
}

/**
 * Checks if the daily appointment limit for the given lawyer has been reached.
 *
 * @param {string} lawyerId - The ID of the lawyer.
 * @param {Date} day - The date to check.
 * @param {Array<microsoftgraph.Event>} allEvents - Array of existing events to check against.
 * @returns {boolean} True if the daily limit has been reached for any day, false otherwise.
 */
function hasDailyLimitConflict(lawyerId, day, allEvents) {
  const lawyer = getLawyer(lawyerId);
  const maxDailyAppointments = lawyer.maxDailyAppointments;

  // Filter events for the specific day
  const eventsForDay = allEvents.filter(event => isSameDay(new Date(event.start.dateTime), day));

  // Filter events for the specific lawyer
  const lawyerEventsForDay = eventsForDay.filter(event => event.categories.includes(lawyerId));

  // Check if the daily limit is reached
  if (lawyerEventsForDay.length >= maxDailyAppointments) {
    // log the conflict
    console.warn(`Daily limit reached for lawyer ${lawyerId} on`, day);
    return true;
  }

  return false;
}

/**
 * Determines if a slot is a virtual meeting.
 * 
 * This function checks the location property of the slot to determine if it includes
 * 'phone' or 'teams', indicating that the meeting is virtual.
 * 
 * @param {{ start: Date, end: Date, location: string }} slot - The slot containing location details.
 * @returns {boolean} True if the slot is virtual, false otherwise.
 */
function isVirtualMeeting(slot) {
  const location = slot.location || "";
  if (!location) {
    console.error("Slot location is missing or invalid for virtual meeting check.");
    return false;
  }

  const isVirtual = [
    'phone',
    'tel',
    'telephone',
    'téléphone',
    'teams',
    'ms teams',
    'microsoft teams',
    'microsoft teams meeting',
  ].some(keyword => location.includes(keyword));
  
  console.log(`Slot location: ${location}, is virtual: ${isVirtual}`);
  return isVirtual;
}

/**
 * Checks if two slots overlap based on their start and end times.
 * @param {{ start: Date, end: Date }} slotA
 * @param {{ start: Date, end: Date }} slotB 
 * @returns {boolean} True if they do, false otherwise. 
 */
export function isOverlapping(slotA, slotB) {
  return slotA.start < slotB.end && slotB.start < slotA.end;
}

/**
 * Checks if two dates are the same day (ignoring time of day).
 * @param {Date} dateA - The first date.
 * @param {Date} dateB - The second date.
 * @returns {boolean} - True if the dates are the same day, false otherwise.
 */
export function isSameDay(dateA, dateB) {
  return dateA.getFullYear() === dateB.getFullYear() &&
         dateA.getMonth() === dateB.getMonth() &&
         dateA.getDate() === dateB.getDate();
}

/**
 * Checks if a proposed time slot for a lawyer is valid (i.e., has no conflicts)
 * by checking for conflicts with existing events.
 *
 * @param {string} lawyerId - The ID of the lawyer
 * @param {{ start: Date, end: Date, location: string }} proposedSlot - The proposed time slot
 * @param {Array<microsoftgraph.Event>} allEvents - Array of existing events to check for conflicts
 * @returns {boolean} True if the proposed time slot is valid, false otherwise
 */
export function isValidSlot(lawyerId, proposedSlot, allEvents) {
  console.log(`Slot received for validation:`, proposedSlot);
  // === Static Rules (slot-inherent checks) ===

  // Availability rule: lawyer not working this location on this day
  if (hasAvailabilityConflict(lawyerId, proposedSlot)) {
    console.warn("Slot rejected by availability rule", proposedSlot);
    return false;
  }
  // Daily appointment limit reached
  if (hasDailyLimitConflict(lawyerId, proposedSlot.start, allEvents)) {
    console.warn("Slot rejected by daily limit", proposedSlot);
    return false;
  }

  // === Dynamic Rules (require checking against other events) ===

  if (hasOfficeConflict(proposedSlot, allEvents)) {
    console.warn("Slot rejected by office conflict", proposedSlot);
    return false;
  }
  if (hasVirtualConflict(lawyerId, proposedSlot, allEvents)) {
    console.warn("Slot rejected by virtual conflict", proposedSlot);
    return false;
  }
  if (hasBreakConflict(lawyerId, proposedSlot, allEvents)) {
    console.warn("Slot rejected by break conflict", proposedSlot);
    return false;
  }

  // Passed all rules
  return true;
}

/**
 * Generates available appointment slots for a lawyer over a specified number of days,
 * avoiding weekends and considering existing events and lunch breaks. Slots are generated
 * based on the lawyer's working hours and required break times between appointments.
 *
 * @param {Lawyer} lawyer - The lawyer object containing working hours and break details.
 * @param {string} location - The location for the appointment slots.
 * @param {Array<microsoftgraph.Event>} allEvents - Array of existing events to check for conflicts when generating slots.
 * @returns {Array<{ start: Date, end: Date, location: string}>} - An array of available time slots with start and end times.
 */
export function generateSlots(lawyer, location, allEvents) {
  // 0. We assume allEvents already comes in sorted
  // 1. Filter events for the specific lawyer
  const lawyerEvents = allEvents.filter(event =>
    // either by category tag…
    event.categories?.includes(lawyer.id)
    // …or by them showing up in attendees
    || event.attendees?.some(a => a.emailAddress.address === lawyer.email)
  );

  // 2. Map to simple time blocks + keep them sorted
  const events = lawyerEvents
    .map(ev => ({
      start: new Date(ev.start.dateTime),
      end:   new Date(ev.end.dateTime),
      raw:   ev
    }))
    .sort((a, b) => a.start - b.start);

  const slots = [];
  const now = new Date();
  const slotDuration   = 60 * 60 * 1000;            // 1 hr
  const requiredBreak  = lawyer.breakMinutes * 60e3; // break in ms

  for (let day = 0; day < RANGE_IN_DAYS; day++) {
    const currentDay = new Date(now);
    currentDay.setDate(now.getDate() + day);

    // Skip weekends
    const weekday = currentDay.getDay();
    if (weekday === 0 || weekday === 6) continue;

    // Build work window
    const [h1, m1] = lawyer.workingHours.start.split(':').map(Number);
    const [h2, m2] = lawyer.workingHours.end.split(':').map(Number);
    const workStart = new Date(currentDay);
    const workEnd   = new Date(currentDay);
    workStart.setHours(h1, m1, 0, 0);
    workEnd.setHours(h2, m2, 0, 0);

    // 3. Grab only today’s blocks for this lawyer
    const dayBlocks = events
      .filter(b => isSameDay(b.start, currentDay) && b.end > workStart && b.start < workEnd);

    // 4. Sweep for gaps between each block
    let cursor = workStart;
    for (const block of dayBlocks) {
      // any gap before this block?
      if (cursor.getTime() + slotDuration <= block.start.getTime()) {
        let startCursor = new Date(cursor);
        while (startCursor.getTime() + slotDuration <= block.start.getTime()) {
          const endCursor = new Date(startCursor.getTime() + slotDuration);

          // skip lunch
          if (!isOverlapping({ start: startCursor, end: endCursor }, LUNCH_SLOT(currentDay))) {
            slots.push(createSlot(startCursor, endCursor, location));
          }

          startCursor = new Date(endCursor);
        }
      }
      // move cursor past this appointment + break
      cursor = new Date(block.end.getTime() + requiredBreak);
    }

    // 5. Final tail gap after last event
    while (cursor.getTime() + slotDuration <= workEnd.getTime()) {
      const endCursor = new Date(cursor.getTime() + slotDuration);

      if (!isOverlapping({ start: cursor, end: endCursor }, LUNCH_SLOT(currentDay))) {
        slots.push(createSlot(cursor, endCursor, location));
      }

      cursor = new Date(endCursor);
    }
  }

  // Log slots
  console.log("Generated slots:", slots);

  return slots;
}