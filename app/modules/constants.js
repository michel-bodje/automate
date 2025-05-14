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
    wordReceiptPage: "word-receipt-page",

    // Button IDs
    backBtn: "back-btn",

    scheduleMenuBtn: "schedule-menu-btn",
    confirmMenuBtn: "conf-menu-btn",
    contractMenuBtn: "contract-menu-btn",
    replyMenuBtn: "reply-menu-btn",
    userManualMenuBtn: "user-manual-btn",
    wordContractMenuBtn: "word-contract-menu-btn",
    wordReceiptMenuBtn: "word-receipt-menu-btn",

    scheduleSubmitBtn: "schedule-appointment-btn",
    confirmSubmitBtn: "send-confirmation-btn",
    contractSubmitBtn: "send-contract-btn",
    replySubmitBtn: "send-reply-btn",
    wordContractSubmitBtn: "generate-word-contract-btn",
    wordReceiptSubmitBtn: "generate-word-receipt-btn",

    // Form input IDs
    scheduleLawyerId: "schedule-lawyer-id",
    confLawyerId: "conf-lawyer-id",
    contractLawyerId: "contract-lawyer-id",
    replyLawyerId: "reply-lawyer-id",
    wordReceiptLawyerId: "receipt-lawyer-id",

    scheduleLocation: "schedule-location",
    confLocation: "conf-location",

    scheduleClientName: "schedule-client-name",
    scheduleClientPhone: "schedule-client-phone",
    wordContractClientName: "word-client-name",
    wordReceiptClientName: "receipt-client-name",

    scheduleClientEmail: "schedule-client-email",
    confClientEmail: "conf-client-email",
    contractClientEmail: "contract-client-email",
    replyClientEmail: "reply-client-email",
    wordContractClientEmail: "word-client-email",

    scheduleClientLanguage: "schedule-client-language",
    confClientLanguage: "conf-client-language",
    contractClientLanguage: "contract-client-language",
    replyClientLanguage: "reply-client-language",
    wordContractClientLanguage: "word-client-language",
    wordReceiptClientLanguage: "receipt-client-language",

    confDate: "conf-date",
    confTime: "conf-time",
    manualDate: "manual-date",
    manualTime: "manual-time",
    scheduleMode: "schedule-mode",
    
    emailContractDeposit: "email-contract-deposit",
    wordContractDeposit: "word-contract-deposit",
    wordReceiptDeposit: "word-receipt-deposit",
    wordContractTitle: "word-contract-title",
    customContractTitle: "custom-contract-title",

    refBarreau: "ref-barreau",
    existingClient: "existing-client",

    scheduleFirstConsultation: "schedule-first-consultation",
    confFirstConsultation: "conf-first-consultation",

    paymentMade: "payment-made",
    schedulePaymentMethod: "payment-method",
    receiptPaymentMethod: "receipt-payment-method",
    paymentOptionsContainer: "payment-options-container",

    // Case details
    caseType: "case-type",
    caseDetails: "case-details",

    spouseName: "spouse-name",
    deceasedName: "deceased-name",
    executorName: "executor-name",
    employerName: "employer-name",
    businessName: "business-name",
    mandateDetails: "mandate-details",
    otherPartyName: "other-party-name",
    commonField: "common-field",

    conflictSearchDoneDivorce: "conflict-search-done-divorce",
    conflictSearchDoneEstate: "conflict-search-done-estate",

    // Notes
    notes: "schedule-notes",
    notesContainer: "notes-container",

    // Misc
    loadingOverlay: "loading-overlay",
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

/**
 * Time-related constants
 */

export const FIRM_TIMEZONE = "America/Toronto";
export const LUNCH_START_HOUR = 13; // 1pm
export const LUNCH_END_HOUR = 14;   // 2pm
export const RANGE_IN_DAYS = 14;    // Centralized time range: 2 weeks

/**
 * Generates the lunch slot for a given day.
 * @param {Date} day - The date for which to generate the lunch slot.
 * @returns {{ start: Date, end: Date }} - The lunch slot with start and end times.
 */
export function LUNCH_SLOT(day) {
const lunchStart = new Date(day.getFullYear(), day.getMonth(), day.getDate(), LUNCH_START_HOUR, 0, 0, 0);
const lunchEnd = new Date(day.getFullYear(), day.getMonth(), day.getDate(), LUNCH_END_HOUR, 0, 0, 0);
return { start: lunchStart, end: lunchEnd };
}