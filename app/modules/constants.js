/**
 * @module constants
 * @description This module contains constants used throughout the application,
 * such as HTML element IDs for various components, including pages, buttons,
 * form inputs, and dropdowns.
 * This avoids hardcoding ID strings in multiple places and makes it easier to maintain the code.
 * This way, if you need to change an ID, you only have to do it here.
 */
export const ELEMENT_IDS = {
    // Page IDs
    mainPage: "menu-page",
    schedulePage: "schedule-page",
    confirmPage: "confirmation-page",
    contractPage: "contract-page",
    replyPage: "reply-page",
    wordContractPage: "word-contract-page",

    // Button IDs
    backBtn: "back-btn",

    scheduleMenuBtn: "schedule-menu-btn",
    confirmMenuBtn: "conf-menu-btn",
    contractMenuBtn: "contract-menu-btn",
    replyMenuBtn: "reply-menu-btn",
    userManualMenuBtn: "user-manual-btn",
    wordContractMenuBtn: "word-contract-menu-btn",

    scheduleSubmitBtn: "schedule-appointment-btn",
    confirmSubmitBtn: "send-confirmation-btn",
    contractSubmitBtn: "send-contract-btn",
    replySubmitBtn: "send-reply-btn",
    wordContractSubmitBtn: "generate-word-contract-btn",

    // Form input IDs
    scheduleLawyerId: "schedule-lawyer-id",
    confLawyerId: "conf-lawyer-id",
    contractLawyerId: "contract-lawyer-id",
    replyLawyerId: "reply-lawyer-id",

    scheduleLocation: "schedule-location",
    confLocation: "conf-location",

    scheduleClientName: "schedule-client-name",
    scheduleClientPhone: "schedule-client-phone",
    wordClientName: "word-client-name",

    scheduleClientEmail: "schedule-client-email",
    confClientEmail: "conf-client-email",
    contractClientEmail: "contract-client-email",
    replyClientEmail: "reply-client-email",
    wordClientEmail: "word-client-email",

    scheduleClientLanguage: "schedule-client-language",
    confClientLanguage: "conf-client-language",
    contractClientLanguage: "contract-client-language",
    replyClientLanguage: "reply-client-language",
    wordClientLanguage: "word-client-language",

    confDate: "conf-date",
    confTime: "conf-time",
    scheduleMode: "schedule-mode",
    manualDate: "manual-date",
    manualTime: "manual-time",
    
    emailContractDeposit: "email-contract-deposit",
    wordContractDeposit: "word-contract-deposit",
    wordContractTitle: "word-contract-title",
    customContractTitle: "custom-contract-title",

    firstConsultation: "first-consultation",
    refBarreau: "ref-barreau",

    paymentMade: "payment-made",
    paymentMethod: "payment-method",
    paymentOptionsContainer: "payment-options-container",

    // Case details
    caseType: "case-type",
    caseDetails: "case-details",

    spouseName: "spouse-name",
    deceasedName: "deceased-name",
    executorName: "executor-name",
    employerName: "employer-name",
    businessName: "business-name",
    mandateDetails: "mandatate-details",
    otherPartyName: "other-party-name",
    commonField: "common-field",

    conflictSearchDoneDivorce: "conflict-search-done-divorce",
    conflictSearchDoneEstate: "conflict-search-done-estate",

    // Notes
    notes: "schedule-notes",
    notesContainer: "notes-container",

    // Misc
    loadingOverlay: "loading-overlay",
    errorMessage: "error-message",
};

/**
 * @constant MS
 * @description Constants related to the Microsoft Authentication Library (MSAL).
 * It includes the client ID, tenant ID, and URL for the application.
 */
export const MS = {
    clientId: "768ccbfe-c251-4fe4-bfeb-eff27fdd356e",
    tenantId: "5ef3243e-fce7-454b-8508-f38fa4259f55",
    urlDev: "https://localhost:3000/taskpane.html",
    urlProd: "https://michel-bodje.github.io/automate/taskpane.html",
};