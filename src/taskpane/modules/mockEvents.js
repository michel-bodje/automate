import { 
    getLawyer,
    lawyers as Lawyers,
    overlapsLunch,
    adjustForLunch
} from "../index.js";

/**
 * Generates an array of mock events for testing and development purposes.
 * The events are generated over the given number of days, with a random number of events per day.
 * Each event has a unique subject line, a random lawyer, a random time slot between 9am and 5pm, and a random location.
 * The time slot is adjusted if needed to avoid the lunch break.
 * @param {number} [daysToGenerate=14] - The number of days to generate events for.
 * @returns {Array} - An array of mock events in Microsoft Graph format.
 */
export function generateMockEvents(daysToGenerate = 14) {
    try {
        const events = [];
        const now = new Date();
        const lawyers = Lawyers.map((lawyer) => lawyer.name);
        const slotDuration = 30 * 60 * 1000; // 30 minutes base slot
        let appointmentCounter = 1;
        const caseTypes = [
            "Divorce Consultation",
            "Estate Planning",
            "Employment Dispute",
            "Contract Review",
            "Real Estate Closing",
            "Name Change",
            "Adoption",
            "Business Incorporation"
        ];

        // Generate events for each day
        for (let day = 0; day < daysToGenerate; day++) {
            const date = new Date(now);
            date.setDate(now.getDate() + day);
            
            // Skip weekends
            if (date.getDay() === 0 || date.getDay() === 6) {
                continue;
            }
      
            // Create 0-3 random events per day (leaving gaps)
            const eventCount = Math.floor(Math.random() * 4);
      
            for (let i = 0; i < eventCount; i++) {
                let startHour = 9 + Math.floor(Math.random() * 8); // 9am-5pm
                let startMin = Math.random() > 0.5 ? 0 : 30; // :00 or :30
                let duration = Math.random() > 0.5 ? 30 : 60; // 30 or 60 mins

                // Create initial time slot
                const initialStart = new Date(date);
                initialStart.setHours(startHour, startMin, 0, 0);
                const initialEnd = new Date(initialStart.getTime() + duration * 60 * 1000);

                // Adjust for lunch break if needed
                const adjustedSlot = adjustForLunch(initialStart, initialEnd, duration * 60 * 1000);
                if (!adjustedSlot) continue; // Skip if can't fit after lunch

                // Generate unique subject line
                const caseType = caseTypes[Math.floor(Math.random() * caseTypes.length)];
                const clientInitial = String.fromCharCode(65 + Math.floor(Math.random() * 26)); // A-Z
                const subject = `[#${appointmentCounter++}] ${caseType} - ${clientInitial}. Client`;

                events.push({
                    subject: subject,
                    start: { 
                        dateTime: adjustedSlot.start,
                    },
                    end: { 
                        dateTime: adjustedSlot.end, 
                    },
                    categories: [lawyers[Math.floor(Math.random() * lawyers.length)]],
                    location: { 
                        displayName: ["Office", "Teams", "Phone"][Math.floor(Math.random() * 3)] 
                    }
                });
            }
        }
        return events;
    } catch (error) {
        console.error("Error generating mock events:", error);
        return [];
    }
}

/**
 * Generates an array of mock test events for controlled testing scenarios
 * and realistic random events. The events are generated over a specified
 * number of days, with controlled scenarios for conflicts and limits, and
 * random events following typical working hours and locations.
 *
 * Controlled scenarios include:
 * - Office Conflict: Two events scheduled at the same time and location.
 * - Virtual Conflict: Overlapping virtual meetings.
 * - Daily Limit Reached: Maximum daily appointments for a lawyer.
 * - Break Time Violation: Insufficient break between appointments.
 * - Lunch Break Intrusion: Event overlapping with lunch time.
 *
 * Random events are generated for remaining days, avoiding weekends,
 * with random times, durations, and locations, adjusted to avoid lunch breaks.
 *
 * @param {number} [daysToGenerate=14] - The number of days to generate events for.
 * @returns {Array} - An array of mock events in Microsoft Graph format.
 */
export function generateMockTestEvents(daysToGenerate = 14) {
    const events = [];
    const now = new Date();
    now.setHours(0, 0, 0, 0); // Start of today

    // Get specific lawyers for testing
    const DH = getLawyer("DH"); // Dorin Holban
    const TG = getLawyer("TG"); // Tim Gagin
    const MM = getLawyer("MM"); // Marie Madelin

    // ========================
    // Controlled Test Scenarios
    // ========================
    
    // Test 1: Office Conflict (Same time, same location)
    const officeConflictDate = new Date(now);
    officeConflictDate.setDate(now.getDate() + 1);
    events.push({
        subject: "TEST: Office Conflict 1",
        start: { dateTime: new Date(officeConflictDate.setHours(10, 0)) },
        end: { dateTime: new Date(officeConflictDate.setHours(11, 0)) },
        categories: [DH.name],
        location: { displayName: "Office" }
    }, {
        subject: "TEST: Office Conflict 2",
        start: { dateTime: new Date(officeConflictDate.setHours(10, 0)) },
        end: { dateTime: new Date(officeConflictDate.setHours(11, 0)) },
        categories: [MM.name],
        location: { displayName: "Office" }
    });

    // Test 2: Virtual Conflict (DH/TG overlap)
    const virtualConflictDate = new Date(now);
    virtualConflictDate.setDate(now.getDate() + 1);
    events.push({
        subject: "TEST: Virtual Conflict 1 (DH)",
        start: { dateTime: new Date(virtualConflictDate.setHours(14, 0)) },
        end: { dateTime: new Date(virtualConflictDate.setHours(15, 0)) },
        categories: [DH.id],
        location: { displayName: "Teams" }
    }, {
        subject: "TEST: Virtual Conflict 2 (TG)",
        start: { dateTime: new Date(virtualConflictDate.setHours(14, 30)) },
        end: { dateTime: new Date(virtualConflictDate.setHours(15, 30)) },
        categories: [TG.id],
        location: { displayName: "Phone" }
    });

    // Test 3: Daily Limit Reached
    const limitTestDate = new Date(now);
    limitTestDate.setDate(now.getDate() + 2);
    for (let i = 0; i < DH.maxDailyAppointments; i++) {
        events.push({
            subject: `TEST: Daily Limit ${i + 1}`,
            start: { dateTime: new Date(limitTestDate.setHours(10 + i, 0)) },
            end: { dateTime: new Date(limitTestDate.setHours(11 + i, 0)) },
            categories: [DH.id],
            location: { displayName: "Office" }
        });
    }

    // Test 4: Break Time Violation
    const breakTestDate = new Date(now);
    breakTestDate.setDate(now.getDate() + 3);
    events.push({
        subject: "TEST: Break Conflict 1",
        start: { dateTime: new Date(breakTestDate.setHours(10, 0)) },
        end: { dateTime: new Date(breakTestDate.setHours(11, 0)) },
        categories: [TG.id],
        location: { displayName: "Teams" }
    }, {
        subject: "TEST: Break Conflict 2",
        start: { dateTime: new Date(breakTestDate.setHours(11, 5)) }, // Only 5min break
        end: { dateTime: new Date(breakTestDate.setHours(12, 5)) },
        categories: [TG.id],
        location: { displayName: "Office" }
    });

    // Test 5: Lunch Break Intrusion
    const lunchTestDate = new Date(now);
    lunchTestDate.setDate(now.getDate() + 4);
    events.push({
        subject: "TEST: Lunch Intrusion",
        start: { dateTime: new Date(lunchTestDate.setHours(12, 30)) },
        end: { dateTime: new Date(lunchTestDate.setHours(13, 30)) },
        categories: [MM.name],
        location: { displayName: "Office" }
    });

    // ======================
    // Realistic Random Events
    // ======================
    for (let day = 5; day < daysToGenerate; day++) {
        const date = new Date(now);
        date.setDate(now.getDate() + day);
        if (date.getDay() === 0 || date.getDay() === 6) continue;

        const eventCount = Math.floor(Math.random() * 3);
        for (let i = 0; i < eventCount; i++) {
            const lawyer = [DH, TG, MM][Math.floor(Math.random() * 3)];
            const startHour = 9 + Math.floor(Math.random() * 8);
            const duration = Math.random() > 0.5 ? 60 : 30;
            const location = ["Office", "Teams", "Phone"][Math.floor(Math.random() * 3)];

            const start = new Date(date);
            start.setHours(startHour, 0, 0, 0);
            const end = new Date(start.getTime() + duration * 60000);

            const adjusted = adjustForLunch(start, end, duration * 60000);
            if (!adjusted) continue;

            events.push({
                subject: `Random Event Day${day}-${i}`,
                start: { dateTime: adjusted.start },
                end: { dateTime: adjusted.end },
                categories: [lawyer.id],
                location: { displayName: location }
            });
        }
    }

    console.log("Generated Test Events:");
    events.forEach(event => {
        if (event.subject.startsWith("TEST:")) {
            console.log(`- ${event.subject.padEnd(30)} ${event.start.dateTime.toLocaleTimeString()} @ ${event.location.displayName}`);
        }
    });

    return events;
}