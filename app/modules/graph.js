import { Client } from '@microsoft/microsoft-graph-client';
import { 
    msalInstance,
    FIRM_TIMEZONE,
    RANGE_IN_DAYS,
    generateMockEvents,
} from '../index.js';

// Create authentication provider
const authProvider = {
    /**
     * Acquires an access token silently if possible, otherwise falls back to interactive
     * acquisition. The token is suitable for accessing Microsoft Graph APIs.
     *
     * @returns {Promise<string>} The acquired access token.
     */
    getAccessToken: async () => {
        try {
            // 1. Clear cache to prevent stale tokens
            // msalInstance.clearCache();

            // 2. Get active account with authority hint
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                const result = await msalInstance.acquireTokenSilent({
                    scopes: ["Calendars.ReadWrite"],
                    account: accounts[0],
                });
                return result.accessToken;
            }

            // 3. Interactive login with explicit tenant
            const login = await msalInstance.loginPopup({
                scopes: ["Calendars.ReadWrite"],
                prompt: "select_account"
            });
            return login.accessToken;
        } catch (error) {
            console.error("Token Debug:", {
                errorCode: error.errorCode,
                message: error.message,
                stack: error.stack
            });
            throw error;
        }
    }
};

// Create Graph client
const client = Client.initWithMiddleware({ authProvider });

/**
 * Using MS Graph, fetches a list of upcoming events from the account calendar,
 * for the standard time range.
 * The events are ordered by start time.
 * 
 * Generates mock events if used in development mode.
 * @returns {Promise<Array<microsoftgraph.Event>>} - A promise resolving to an array of events.
 */
export async function fetchCalendarEvents() {
    if (process.env.NODE_ENV === "development") {
        console.warn("Using mock calendar data for testing");
        return generateMockEvents();
    }

    try {
        const now = new Date();
        const range = {
            start: now,
            end: new Date(now.setDate(now.getDate() + RANGE_IN_DAYS))
        };
        const events = await client
            .api('/me/calendarView')
            .header('Prefer', `outlook.timezone="${FIRM_TIMEZONE}"`)
            .query({
                startDateTime: range.start.toISOString(),
                endDateTime: range.end.toISOString(),
                $select: 'subject,start,end,location,attendees,categories',
                $expand: 'instances',
                $orderby: 'start/dateTime',
                $top: 99,
            })
            .get();
        
        console.log("Fetched events:", events.value.map(event => ({
            subject: event.subject,
            start: event.start.dateTime,
            end: event.end.dateTime,
            location: event.location.displayName,
            attendees: event.attendees.map(attendee => ({
                name: attendee.emailAddress.name,
            })),
            categories: event.categories,
        })));
        return events.value;
    } catch (error) {
        console.error('Graph API Error:', {
            error,
            range,
        });
        throw error;
    }
}