import { PublicClientApplication } from "@azure/msal-browser";

const msalConfig = {
    auth: {
        clientId: "a948091c-c2dd-42a8-9e27-c2092740ab74",
        authority: "https://login.microsoftonline.com/7d3dbdb0-f9e1-4027-9481-ea91a040f43b",
        redirectUri: "https://localhost:3000/taskpane.html",
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