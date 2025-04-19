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
  return slotA.start <= slotB.end && slotB.start <= slotA.end;
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
  // Check the proposed slot against each rule
  if (hasAvailabilityConflict(lawyerId, proposedSlot)) {
    console.warn(`Slot rejected due to availability conflict:`, proposedSlot);
    return false;
  }
  if (hasDailyLimitConflict(lawyerId, proposedSlot.start, allEvents)) {
    console.warn(`Slot rejected due to daily limit conflict:`, proposedSlot);
    return false;
  }
  if (hasOfficeConflict(proposedSlot, allEvents)) {
    console.warn(`Slot rejected due to office conflict:`, proposedSlot);
    return false;
  }
  if (hasVirtualConflict(lawyerId, proposedSlot, allEvents)) {
    console.warn(`Slot rejected due to virtual conflict:`, proposedSlot);
    return false;
  }
  if (hasBreakConflict(lawyerId, proposedSlot, allEvents)) {
    console.warn(`Slot rejected due to break conflict:`, proposedSlot);
    return false;
  }

  console.log(`Slot is valid:`, proposedSlot);
  return true; // Slot is valid if no conflicts are found
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
  const now = new Date();
  const slots = [];
  const slotDuration = 60 * 60 * 1000; // 1 hour in milliseconds
  const requiredBreak = lawyer.breakMinutes * 60 * 1000; // Break time in milliseconds

  for (let day = 0; day < RANGE_IN_DAYS; day++) {
    const currentDay = now;
    currentDay.setDate(now.getDate() + day);

    // Skip weekends
    if (currentDay.getDay() === 0 || currentDay.getDay() === 6) continue;

    // Skip the day if the daily limit is reached
    if (hasDailyLimitConflict(lawyer.id, currentDay, allEvents)) {
      console.warn(`Skipping day ${currentDay.toDateString()} due to daily limit conflict.`);
      continue;
    }

    // Define working hours for the day
    const [startHour, startMin] = lawyer.workingHours.start.split(":").map(Number);
    const [endHour, endMin] = lawyer.workingHours.end.split(":").map(Number);

    const workStart = new Date(currentDay);
    workStart.setHours(startHour, startMin, 0, 0);

    const workEnd = new Date(currentDay);
    workEnd.setHours(endHour, endMin, 0, 0);

    // Filter and sort events for the current day
    const dayEvents = allEvents
      .filter(event => 
      (event.categories?.includes(lawyer.name) || 
       event.attendees?.some(attendee => attendee.emailAddress?.name?.toLowerCase().includes(lawyer.name.toLowerCase()))) &&
      isSameDay(new Date(event.start.dateTime), currentDay)
      )
      .sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime));

    let lastEventEnd = workStart;

    // Generate slots between events
    for (const event of dayEvents) {
      const eventStart = new Date(event.start.dateTime);
      const eventEnd = new Date(event.end.dateTime);

      if (lastEventEnd < eventStart) {
        let potentialSlotStart = new Date(lastEventEnd);
        while (potentialSlotStart.getTime() + slotDuration <= eventStart.getTime()) {
          const potentialSlotEnd = new Date(potentialSlotStart.getTime() + slotDuration);

          // Skip slot if it overlaps lunch
          if (!isOverlapping({ start: potentialSlotStart, end: potentialSlotEnd }, LUNCH_SLOT(currentDay))) {
          slots.push(createSlot(potentialSlotStart, potentialSlotEnd, location));
          }

          potentialSlotStart = new Date(potentialSlotEnd);
        }
      }

      lastEventEnd = new Date(eventEnd.getTime() + requiredBreak);
    }

    // Generate slots after the last event of the day
    let potentialSlotStart = new Date(lastEventEnd);
    while (potentialSlotStart.getTime() + slotDuration <= workEnd.getTime()) {
      const potentialSlotEnd = new Date(potentialSlotStart.getTime() + slotDuration);

      // Skip slot if it overlaps lunch
      if (!isOverlapping({ start: potentialSlotStart, end: potentialSlotEnd }, LUNCH_SLOT(currentDay))) {
        slots.push(createSlot(potentialSlotStart, potentialSlotEnd, location));
      }

      potentialSlotStart = new Date(potentialSlotEnd);
    }
  }
  console.log("Generated slots:", slots);
  return slots;
}