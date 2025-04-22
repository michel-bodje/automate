import { MS } from "./constants.js";
import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
    auth: {
        clientId: `${MS.clientId}`,
        authority: `https://login.microsoftonline.com/${MS.tenantId}`,
        redirectUri: `${MS.urlProd}`,
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