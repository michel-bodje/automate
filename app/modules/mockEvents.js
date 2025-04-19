import { getLawyer, overlapsLunch } from "../index.js";

/**
 * Generates an array of mock events for testing and development purposes.
 * The events include controlled test scenarios and realistic random events.
 *
 * @param {number} [daysToGenerate=14] - The number of days to generate events for.
 * @returns {Array} - An array of mock events in Microsoft Graph format.
 */
export function generateMockEvents(daysToGenerate = 14) {
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

    // Test 2: Virtual Conflict (Overlapping virtual meetings)
    const virtualConflictDate = new Date(now);
    virtualConflictDate.setDate(now.getDate() + 2);
    events.push({
        subject: "TEST: Virtual Conflict 1 (DH)",
        start: { dateTime: new Date(virtualConflictDate.setHours(14, 0)) },
        end: { dateTime: new Date(virtualConflictDate.setHours(15, 0)) },
        categories: [DH.name],
        location: { displayName: "Teams" }
    }, {
        subject: "TEST: Virtual Conflict 2 (TG)",
        start: { dateTime: new Date(virtualConflictDate.setHours(14, 30)) },
        end: { dateTime: new Date(virtualConflictDate.setHours(15, 30)) },
        categories: [TG.name],
        location: { displayName: "Phone" }
    });

    // Test 3: Daily Limit Reached
    const dailyLimitDate = new Date(now);
    dailyLimitDate.setDate(now.getDate() + 3);
    for (let i = 0; i < DH.maxDailyAppointments; i++) {
        events.push({
            subject: `TEST: Daily Limit ${i + 1}`,
            start: { dateTime: new Date(dailyLimitDate.setHours(9 + i, 0)) },
            end: { dateTime: new Date(dailyLimitDate.setHours(10 + i, 0)) },
            categories: [DH.name],
            location: { displayName: "Office" }
        });
    }

    // Test 4: Break Time Violation
    const breakConflictDate = new Date(now);
    breakConflictDate.setDate(now.getDate() + 4);
    events.push({
        subject: "TEST: Break Conflict 1",
        start: { dateTime: new Date(breakConflictDate.setHours(10, 0)) },
        end: { dateTime: new Date(breakConflictDate.setHours(11, 0)) },
        categories: [TG.name],
        location: { displayName: "Teams" }
    }, {
        subject: "TEST: Break Conflict 2",
        start: { dateTime: new Date(breakConflictDate.setHours(11, 5)) }, // Only 5 min break
        end: { dateTime: new Date(breakConflictDate.setHours(12, 5)) },
        categories: [TG.name],
        location: { displayName: "Office" }
    });

    // Test 5: Lunch Break Intrusion
    const lunchConflictDate = new Date(now);
    lunchConflictDate.setDate(now.getDate() + 5);
    events.push({
        subject: "TEST: Lunch Intrusion",
        start: { dateTime: new Date(lunchConflictDate.setHours(12, 30)) },
        end: { dateTime: new Date(lunchConflictDate.setHours(13, 30)) },
        categories: [MM.name],
        location: { displayName: "Office" }
    });

    // ======================
    // Realistic Random Events
    // ======================
    for (let day = 6; day < daysToGenerate; day++) {
        const date = new Date(now);
        date.setDate(now.getDate() + day);

        // Skip weekends
        if (date.getDay() === 0 || date.getDay() === 6) continue;

        const eventCount = Math.floor(Math.random() * 3); // 0-2 events per day
        for (let i = 0; i < eventCount; i++) {
            const lawyer = [DH, TG, MM][Math.floor(Math.random() * 3)];
            const startHour = 9 + Math.floor(Math.random() * 8);
            const duration = Math.random() > 0.5 ? 60 : 30;
            const location = ["Office", "Teams", "Phone"][Math.floor(Math.random() * 3)];

            const start = new Date(date);
            start.setHours(startHour, 0, 0, 0);
            const end = new Date(start.getTime() + duration * 60000);

            // Skip events that overlap lunch (12:00 - 13:00)
            if (overlapsLunch(start, end)) continue;

            events.push({
                subject: `Random Event Day${day}-${i}`,
                start: { dateTime: start },
                end: { dateTime: end },
                categories: [lawyer.name],
                location: { displayName: location }
            });
        }
    }

    console.log("Generated Test Events:", events);
    return events;
}