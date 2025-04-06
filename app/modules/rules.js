import { 
  Lawyer,
  getLawyer,
  overlapsLunch,
  adjustForLunch
} from "../index.js";

export const locationRules = {
  // Centralized list of locations
  locations: ["Office", "Phone", "Teams"],
  
  // List of unavailability for each lawyer
  lawyerUnavailability: {
    // TODO: Implement location restrictions
  },
};

/**
 * Get available locations for a lawyer based on the current day.
 * @param {string} lawyerId - The ID of the lawyer.
 * @returns {Array} - List of available locations for the lawyer.
 */
export function getAvailableLocations(lawyerId) {
  // Get current day
  const today = new Date().toLocaleString("en-US", { weekday: "long" });

  const unavailability = locationRules.lawyerUnavailability?.lawyerId || null;

  // If the lawyer is not predefined, assume full availability
  if (lawyerId in locationRules.lawyerUnavailability === false) {
    return locationRules.locations;
  }

  // Filter locations that are not in the unavailability list for today
  return locationRules.locations.filter((location) => {
    if (!unavailability) {
      return true; // If unavailability is null or undefined, all locations are available
    }
    return !unavailability[location.toLowerCase()]?.includes(today);
  });
}

/**
 * Generates available appointment slots for a lawyer over a specified number of days,
 * avoiding weekends and considering existing events and lunch breaks. Slots are generated
 * based on the lawyer's working hours and required break times between appointments.
 *
 * @param {Array<MicrosoftGraph.Event>} allEvents - Array of existing events to check for conflicts when generating slots.
 * @param {Lawyer} lawyer - The lawyer object containing working hours and break details.
 * @param {string} location - The location for the appointment slots.
 * @param {Date} startDateTime - The start date and time for generating slots.
 * @param {Date} endDateTime - The end date and time for generating slots.
 * @returns {Array<{ start: Date, end: Date, location: string}>} - An array of available time slots with start and end times.
 */
export function generateSlots(allEvents, lawyer, location, startDateTime, endDateTime) {
  const slots = [];
  const slotDuration = 60 * 60 * 1000; // 60 minutes

  const daysToCheck = Math.ceil((endDateTime - startDateTime) / (1000 * 60 * 60 * 24));

  const requiredBreak = lawyer.breakMinutes * 60 * 1000;

  const today = startDateTime;

  for (let day = 0; day < daysToCheck; day++) {
    const dayStart = new Date(today);
    dayStart.setDate(today.getDate() + day);

    // Skip weekends
    if (dayStart.getDay() === 0 || dayStart.getDay() === 6) continue;

    // Parse working hours
    const [startHour, startMin] = lawyer.workingHours.start.split(":").map(Number);
    const [endHour, endMin] = lawyer.workingHours.end.split(":").map(Number);
    
    const workStart = new Date(dayStart);
    workStart.setHours(startHour, startMin, 0, 0);
    
    const workEnd = new Date(dayStart);
    workEnd.setHours(endHour, endMin, 0, 0);

    // Get lawyer's events for this day
    const lawyerEvents = allEvents
      .filter(event => 
        event.categories?.includes(lawyer.name) &&
        isSameDay(new Date(event.start.dateTime), dayStart)
      )
      .sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime));
    
    let lastEventEnd = workStart;

    // Generate slots between existing events
    for (const event of lawyerEvents) {
      const eventStart = new Date(event.start.dateTime);
      const eventEnd = new Date(event.end.dateTime);
      
      if (lastEventEnd < eventStart) {
        const potentialSlotEnd = new Date(lastEventEnd.getTime() + slotDuration);
        const adjustedSlot = adjustForLunch(lastEventEnd, potentialSlotEnd, slotDuration);
        
        if (adjustedSlot && adjustedSlot.end <= eventStart) {
          slots.push(createSlot(adjustedSlot.start, adjustedSlot.end, location));
        }
      }

      lastEventEnd = new Date(eventEnd.getTime() + requiredBreak);
    }

    // Final slot of the day
    if (lastEventEnd < workEnd) {
      const potentialSlotEnd = new Date(lastEventEnd.getTime() + slotDuration);
      const adjustedSlot = adjustForLunch(lastEventEnd, potentialSlotEnd, slotDuration);
      
      if (adjustedSlot && adjustedSlot.end <= workEnd) {
        slots.push(createSlot(adjustedSlot.start, adjustedSlot.end, location));
      }
    }
  }

  console.log("Generated slots:", slots.map(s => ({
    start: s.start,
    end: s.end,
    location: s.location
  })));

  return slots;
}

/**
 * Checks if a proposed time slot for a lawyer is valid (i.e., has no conflicts)
 * by checking for conflicts with existing events.
 *
 * @param {string} lawyerId - The ID of the lawyer
 * @param {{ start: Date, end: Date, location: string }} proposedSlot - The proposed time slot
 * @param {Array<MicrosoftGraph.Event>} allEvents - Array of existing events to check for conflicts
 * @returns {boolean} - true if the proposed time slot is valid, false otherwise
 */
export function isValidSlot(lawyerId, proposedSlot, allEvents) {
  const proposedEvent = {
    start: { dateTime: proposedSlot.start.toString() },
    end: { dateTime: proposedSlot.end.toString() },
    location: { displayName: proposedSlot.location },
    categories: [lawyerId],
  };
  
  // If any of the conflicts is true, the proposed slot is invalid
  return [
    hasOfficeConflict(proposedEvent, allEvents),
    hasVirtualConflict(lawyerId, proposedEvent, allEvents),
    hasDailyLimitConflict(lawyerId, allEvents),
    hasBreakConflict(lawyerId, proposedEvent, allEvents),
  ].every((conflict) => !conflict);
}

/**
 * Checks if a proposed event has a location conflict with existing events.
 * A location conflict occurs when two events are scheduled in the office
 * at the same time.
 *
 * @param {Object} proposedEvent - The event to check, with location and start/end times.
 * @param {Array<MicrosoftGraph.Event>} allEvents - Array of existing events to check for conflicts.
 * @returns {boolean} - true if there is a location conflict, false otherwise.
 */
function hasOfficeConflict(proposedEvent, allEvents) {
  try {
    const proposedLocation = proposedEvent.location.displayName?.toLowerCase();

    // If the proposed event is not in the office, return false
    if (!proposedLocation === "office") {
      return false;
    }

    // Check if any existing event is scheduled in the office at the same time
    for (const existingEvent of allEvents) {
      const existingLocation = existingEvent.location?.displayName?.toLowerCase();
      if (existingLocation === "office" && isOverlapping(proposedEvent, existingEvent)) {
        console.warn(`Office location conflict: ${proposedEvent.start.dateTime} with ${existingEvent.subject}`);
        return true;
      }
    }
    return false;
  } catch (error) {
    console.error("Error checking for office location conflict:", error);
    // on fail, don't block scheduling. You'll have to check manually
    return false;
  }
}

/**
 * Checks if a proposed event has a virtual conflict with existing events.
 * A virtual conflict occurs when two virtual meetings are scheduled with
 * either Dorin Holban or Tim Gagin at the same time.
 *
 * @param {Object} proposedEvent - The event to check, with location and start/end times.
 * @param {Array<MicrosoftGraph.Event>} allEvents - Array of existing events to check for conflicts.
 * @param {string} lawyerId - The ID of the lawyer.
 * @returns {boolean} - true if there is a virtual conflict, false otherwise.
 */
function hasVirtualConflict(lawyerId, proposedEvent, allEvents) {
  try {
    // If the lawyer is not DH or TG, return false
    if (!["DH", "TG"].includes(lawyerId)) {
      // log no conflict
      console.warn(`No virtual conflict for lawyer ${lawyer.name} on ${proposedEvent.start.dateTime}`);
      return false;
    }

    const lawyer = getLawyer(lawyerId);
    const otherLawyerId = lawyerId === "DH" ? "TG" : "DH"; 
    const otherLawyer = getLawyer(otherLawyerId);

    const proposedIsVirtual = isVirtualMeeting(proposedEvent);

    // Check if any existing event is scheduled in a virtual meeting 
    // with the other lawyer at the same time
    let conflict = false;
    for (const existing of allEvents) {
      if (
        existing.categories?.includes(otherLawyer) &&
        isVirtualMeeting(existing) &&
        isOverlapping(proposedEvent, existing) &&
        proposedIsVirtual
      ) {
        console.warn(`Virtual conflict: ${proposedEvent.start.dateTime} with ${existing.subject}`);
        conflict = true;
        break;
      }
    }
    return conflict;
  } catch (error) {
    console.error("Error checking for DH_TG conflict:", error);
    // on fail, don't block scheduling. You'll have to check manually
    return false;
  }
}

/**
 * Checks if the daily appointment limit for the given lawyer has been reached for any day in the range of existing events.
 *
 * @param {string} lawyerId - The ID of the lawyer.
 * @param {{ start: Date, end: Date }} proposedSlot - The proposed appointment slot.
 * @param {Array<MicrosoftGraph.Event>} allEvents - Array of existing events to check against.
 * @returns {boolean} - true if the daily limit has been reached for any day, false otherwise.
 */
function hasDailyLimitConflict(lawyerId, allEvents) {
  const lawyer = getLawyer(lawyerId);
  const maxDailyAppointments = lawyer.maxDailyAppointments;

  const startDate = new Date(allEvents[0].start.dateTime);
  const endDate = new Date(allEvents[allEvents.length - 1].start.dateTime);

  for (let day = startDate; day <= endDate; day.setDate(day.getDate() + 1)) {
    const eventsForDay = allEvents.filter((event) => isSameDay(new Date(event.start.dateTime), day));

    const lawyerEventsForDay = eventsForDay.filter((event) => event.categories.includes(lawyerId));

    if (lawyerEventsForDay.length >= maxDailyAppointments) {
      // log the conflict
      console.warn(`Daily limit reached for lawyer ${lawyerId} on ${day.toISOString()}`);
      return true; // Daily limit reached for this day
    }
  }

  return false; // Daily limit not reached for any day
}

/**
 * Checks if a proposed appointment slot conflicts with the required break time
 * for the given lawyer. A break conflict occurs when there is not enough time
 * between the proposed appointment and either the previous or next appointment
 * with the same lawyer.
 *
 * @param {string} lawyerId - The ID of the lawyer.
 * @param {{ start: Date, end: Date }} proposedSlot - The proposed appointment slot.
 * @param {Array<MicrosoftGraph.Event>} allEvents - Array of existing events to check against.
 * @returns {boolean} - true if there is a break conflict, false otherwise.
 */
function hasBreakConflict(lawyerId, proposedSlot, allEvents) {
  const lawyer = getLawyer(lawyerId);
  const requiredBreak = lawyer.breakMinutes * (60 * 1000);

  // Check previous event
  const previousEvents = allEvents
      .filter(event => event.categories?.includes(lawyerId))
      .sort((a, b) => new Date(b.end.dateTime) - new Date(a.end.dateTime));

  if (previousEvents.length > 0) {
      const lastEventEnd = new Date(previousEvents[0].end.dateTime);
      const breakTime = proposedSlot.start.getTime() - lastEventEnd.getTime();
      if (breakTime < requiredBreak) {
        // log the conflict
        console.warn(`Break conflict: ${proposedSlot.start} with ${previousEvents[0].subject}`);
        return true;
      }
  }

  // Check next event
  const nextEvent = allEvents.find(event => 
      event.categories?.includes(lawyerId) &&
      new Date(event.start.dateTime) > proposedSlot.start
  );

  if (nextEvent) {
      const breakTime = new Date(nextEvent.start.dateTime) - proposedSlot.end.getTime();
      if (breakTime < requiredBreak) {
        // log the conflict
        console.warn(`Break conflict: ${proposedSlot.end} with ${nextEvent.subject}`);
        return true;
      }
  }
  return false;
}

/**
 * Creates a time slot given start and end times.
 * @param {Date} start - The start time of the event.
 * @param {Date} end - The end time of the event.
 * @param {string} location - The location of the event.
 * @returns {{start: Date, end: Date, location: string}} slot - A time slot object with start, end, and location properties.
 */
function createSlot(start, end, location) {
  const slot = {
    start: start,
    end: end,
    location: location
  };
  return slot;
}

/**
 * Determines if an event is a virtual meeting.
 * 
 * This function checks the location of the event to determine if it includes
 * 'phone' or 'teams', indicating that the meeting is virtual.
 * 
 * @param {Object} event - The event object containing location details.
 * @returns {boolean} - True if the event is virtual, false otherwise.
 */
function isVirtualMeeting(event) {
  const location = event.location?.displayName?.toLowerCase() ?? "";
  const isVirtual = location.includes('phone') || location.includes('teams');
  console.log(`Event location: ${location}, is virtual: ${isVirtual}`);
  return isVirtual;
}

/**
 * Checks if two dates are the same day (ignoring time of day).
 * @param {Date} dateA - The first date.
 * @param {Date} dateB - The second date.
 * @returns {boolean} - True if the dates are the same day, false otherwise.
 */
function isSameDay(dateA, dateB) {
  return dateA.getFullYear() === dateB.getFullYear() &&
         dateA.getMonth() === dateB.getMonth() &&
         dateA.getDate() === dateB.getDate();
}

/**
 * Checks if two events overlap based on their start and end times.
 * 
 * @param {Object} eventA - The first event object containing start and end properties.
 * @param {Object} eventB - The second event object containing start and end properties.
 * @returns {boolean} - True if the events overlap, false otherwise.
 */
function isOverlapping(eventA, eventB) {
  const startA = new Date(eventA.start.dateTime);
  const endA = new Date(eventA.end.dateTime);
  const startB = new Date(eventB.start.dateTime);
  const endB = new Date(eventB.end.dateTime);

  return startA <= endB && startB <= endA;
}