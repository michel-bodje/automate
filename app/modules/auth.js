import { PublicClientApplication } from "@azure/msal-browser";
import { MSAL } from "../index.js";

const msalConfig = {
    auth: {
        clientId: `${MSAL.clientId}`,
        authority: `https://login.microsoftonline.com/${MSAL.tenantId}`,
        redirectUri: `${MSAL.url}`,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message) => console.log(`MSAL: ${message}`),
            logLevel: "Verbose"
        }
    }
};

// Initialize MSAL outside of export
export const msalInstance = new PublicClientApplication(msalConfig);