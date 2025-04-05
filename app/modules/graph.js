import { Client } from '@microsoft/microsoft-graph-client';
import { 
    msalInstance,
    FIRM_TIMEZONE,
    generateMockEvents
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
                    forceRefresh: true
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
 * Fetches a list of upcoming events from the given lawyer's calendar, within
 * the given time range. The events are ordered by start time.
 *
 * @param {Lawyer} lawyer - The lawyer to fetch events for.
 * @param {TimeRange} timeRange - The time range to fetch events within.
 * @returns {Promise<Array<MicrosoftGraph.Event>>} - A promise resolving to an array of events.
 */
export async function fetchCalendarEvents(lawyer, timeRange) {
    if (process.env.NODE_ENV === "development") {
        console.warn("Using mock calendar data for testing");
        return generateMockEvents();
    }

    try {
        const events = await client
            .api('/me/calendarView')
            .header('Prefer', `outlook.timezone="${FIRM_TIMEZONE}"`)
            .query({
                startDateTime: timeRange.startDateTime,
                endDateTime: timeRange.endDateTime,
                $select: 'subject,start,end,location,attendees,categories',
                $expand: 'instances',
                $orderby: 'start/dateTime',
                $top: 30,
            })
            .get();
        return events.value;
    } catch (error) {
        console.error('Graph API Error:', {
            message: error.message,
            code: error.code,
            statusCode: error.statusCode,
            timeRange,
            lawyer: lawyer.id
        });
        return [];
    }
}