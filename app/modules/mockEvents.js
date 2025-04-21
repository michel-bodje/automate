import {
  getLawyer,
  isOverlapping,
  LUNCH_SLOT,
  RANGE_IN_DAYS,
} from "../index.js";

/**
 * Generates realistic mock events over the next two weeks for three lawyers.
 * - MM: almost fully booked (4-6 events/day)
 * - DH: moderately booked (2-4 events/day), no office events on Mondays
 * - TG: lightly booked (0-2 events/day), no office events on Fridays
 * Events start on the hour or half-hour, avoid lunch (1–2pm), and do not overlap.
 * @returns {Array<microsoftgraph.Event>}
 */
export function generateMockEvents() {
  const events = [];
  const now = new Date();
  now.setHours(0, 0, 0, 0); // start of today

  // Helper for random integer in [min, max]
  function randInt(min, max) {
    return Math.floor(Math.random() * (max - min + 1)) + min;
  }

  const lawyerIds = ["MM", "DH", "TG"];

  for (let day = 0; day < RANGE_IN_DAYS; day++) {
    const currentDay = new Date(now);
    currentDay.setDate(now.getDate() + day);

    // Skip weekends
    const wd = currentDay.getDay();
    if (wd === 0 || wd === 6) continue;

    for (const id of lawyerIds) {
      const lawyer = getLawyer(id);
      // Determine event count per lawyer
      let count;
      if (id === "MM") count = randInt(4, 6);
      else if (id === "DH") count = randInt(2, 4);
      else /* TG */ count = randInt(0, 2);

      // Prepare working hours
      const [sh, sm] = lawyer.workingHours.start.split(":").map(Number);
      const [eh, em] = lawyer.workingHours.end.split(":").map(Number);
      const workStart = new Date(currentDay);
      const workEnd = new Date(currentDay);
      workStart.setHours(sh, sm, 0, 0);
      workEnd.setHours(eh, em, 0, 0);

      // Generate non-overlapping events
      let attempts = 0;
      let created = 0;
      while (created < count && attempts < count * 10) {
        attempts++;
        // Choose random duration 30 or 60 min
        const duration = 60;
        // Random start hour such that end <= workEnd
        const maxStartHour = eh - Math.ceil(duration / 60);
        const hr = randInt(sh, maxStartHour);
        const mn = randInt(0, 1) === 0 ? 0 : 30;
        const start = new Date(currentDay);
        start.setHours(hr, mn, 0, 0);
        const end = new Date(start.getTime() + duration * 60000);

        // Skip if out of working window
        if (start < workStart || end > workEnd) continue;
        // Skip lunch overlap
        if (isOverlapping({ start, end }, LUNCH_SLOT(currentDay))) continue;

        // Choose location, respecting special rules
        const allLocs = ["Office", "Phone", "Teams"];
        let locs = allLocs.slice();
        if (id === "DH" && wd === 1) locs = locs.filter(l => l !== "Office");
        if (id === "TG" && wd === 5) locs = locs.filter(l => l !== "Office");
        if (!locs.length) locs = allLocs; // fallback
        const location = locs[randInt(0, locs.length - 1)];

        // Check overlap with this lawyer's existing mock events
        const conflict = events.some(ev =>
          ev.categories.includes(lawyer.name) &&
          isOverlapping(
            { start, end },
            { start: new Date(ev.start.dateTime), end: new Date(ev.end.dateTime) }
          )
        );
        if (conflict) continue;

        // Build event in Graph format, include attendees and categories
        events.push({
          subject: `Mock: ${lawyer.name} ${location} ${start.toTimeString().slice(0,5)}`,
          start: { dateTime: start },
          end:   { dateTime: end },
          attendees: [
            {
              emailAddress: {
                name: lawyer.name,
                address: lawyer.email
              },
              type: "required"
            }
          ],
          categories: [lawyer.name],
          location: { displayName: location }
        });
        created++;
      }
    }
  }

  // Sort events by start time
  events.sort((a, b) => new Date(a.start.dateTime) - new Date(b.start.dateTime));

  console.log("Generated Mock Events:", events);
  return events;
}