import "./styles/main.css";
import {
  ELEMENT_IDS,
  formState,
  getLawyer,
  showPage,
  resetPage,
  populateLawyerDropdown,
  populateLanguageDropdown,
  populateLocationDropdown,
  populateCaseTypeDropdown,
  handleCaseDetails,
  handlePaymentOptions,
  isValidInputs,
  isValidSlot,
  generateSlots,
  msalInstance,
  fetchCalendarEvents,
  createEmail,
  createMeeting,
  showLoading,
  showError
} from "./index.js";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // 1. Initialize MSAL for authentication
    await msalInstance.initialize();

    // 2. Populate static UI components
    populateLawyerDropdown();
    populateLanguageDropdown();

    // 3. Setup form interactions
    attachEventListeners();

    // 4. Reset to main menu
    resetPage();
  }
});

/** Attaches all event listeners for the application. */
function attachEventListeners() {
  // Form input event listeners
  const formInputs = document.querySelectorAll('input, select, textarea');

  formInputs.forEach((input) => {
    input.addEventListener('change', (event) => {
      const { id, value } = event.target;
      /**
       * Update form state
       * and activate events based on input id
       */
      console.log("changed: %s", input.id); // debug
      switch (id) {
        case ELEMENT_IDS.scheduleLawyerId:
        case ELEMENT_IDS.confLawyerId:
        case ELEMENT_IDS.contractLawyerId:
        case ELEMENT_IDS.replyLawyerId:
          // lawyer dropdown change
          formState.update("lawyerId", value);
          populateLocationDropdown();
          populateCaseTypeDropdown();
          break;
        case ELEMENT_IDS.scheduleLocation:
        case ELEMENT_IDS.confLocation:
          // location dropdown change
          formState.update("location", value);
          break;
        case ELEMENT_IDS.caseType:
          // case type dropdown change
          formState.update("caseType", value);
          handleCaseDetails();
          break;
        case ELEMENT_IDS.scheduleClientName:
          // client name input change
          formState.update("clientName", value);
          break;
        case ELEMENT_IDS.scheduleClientPhone:
          // client phone input change
          formState.update("clientPhone", value);
          break;
        case ELEMENT_IDS.scheduleClientEmail:
        case ELEMENT_IDS.confClientEmail:
        case ELEMENT_IDS.contractClientEmail:
        case ELEMENT_IDS.replyClientEmail:
          // client email input change
          formState.update("clientEmail", value);
          break;
        case ELEMENT_IDS.scheduleClientLanguage:
        case ELEMENT_IDS.confClientLanguage:
        case ELEMENT_IDS.contractClientLanguage:
        case ELEMENT_IDS.replyClientLanguage:
          // client language dropdown change
          formState.update("clientLanguage", value);
          break;
        case ELEMENT_IDS.confDate:
        case ELEMENT_IDS.confTime:
          // appointment date and time input change
          const dateInput = document.getElementById(ELEMENT_IDS.confDate).value;
          const timeInput = document.getElementById(ELEMENT_IDS.confTime).value;
          if (dateInput && timeInput) {
            const dateTime = new Date(`${dateInput}T${timeInput}`);
            formState.update("appointmentDateTime", dateTime);
          }
          break;
        case ELEMENT_IDS.firstConsultation:
          // first consultation checkbox change
          formState.update("isFirstConsultation", event.target.checked);
          break;
        case ELEMENT_IDS.refBarreau:
          // ref barreau checkbox change
          formState.update("isRefBarreau", event.target.checked);
          break;
        case ELEMENT_IDS.paymentMade:
          // payment checkbox change
          formState.update("isPaymentMade", event.target.checked);
          handlePaymentOptions();
          break;
        case ELEMENT_IDS.paymentMethod:
          // payment method dropdown change
          formState.update("paymentMethod", value);
          break;
        case ELEMENT_IDS.notes:
          // schedule notes textarea change
          formState.update("notes", value);
          break;
        case ELEMENT_IDS.contractDeposit:
          // contract deposit input change
          formState.update("deposit", value);
          break;
        default:
          // no change
          break;
      }
    });
  });

  // Menu buttons event listeners
  const menuButtons = document.querySelectorAll('.menu-btn');

  menuButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      const { id } = event.target;
      switch (id) {
        case ELEMENT_IDS.scheduleMenuBtn:
          showPage(ELEMENT_IDS.schedulePage);
          break;
        case ELEMENT_IDS.confirmMenuBtn:
          showPage(ELEMENT_IDS.confirmPage);
          break;
        case ELEMENT_IDS.contractMenuBtn:
          showPage(ELEMENT_IDS.contractPage);
          break;
        case ELEMENT_IDS.replyMenuBtn:
          showPage(ELEMENT_IDS.replyPage);
          break;
        case ELEMENT_IDS.userManualMenuBtn:
          const manualWindow = window.open('./user-manual.html', '_blank', 'width=684,height=800', false);
          manualWindow.document.title = "User Manual";
          manualWindow.onload = () => {
            const style = manualWindow.document.createElement('style');
            style.textContent = `
              body {
                background-color: #1e1e1e;
                color: #f4f4f4;
                font-family: monospace;
                line-height: 1.6;
                margin: 0;
                padding: 1rem;
              }
              a {
                color: #0078d4;
                text-decoration: none;
              }
              a:hover {
                text-decoration: underline;
              }
            `;
            manualWindow.document.head.appendChild(style);
          };
          break;
        default:
          break;
      }
    });
  });

  // Back buttons event listeners
  const backButtons = document.querySelectorAll('.back-btn');

  backButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      event.preventDefault();
      resetPage();
    });
  });

  // Submit buttons event listeners
  const submitButtons = document.querySelectorAll('.submit-btn');

  submitButtons.forEach((button) => {
    button.addEventListener('click', (event) => {
      event.preventDefault();
      const { id } = event.target;
      switch (id) {
        case ELEMENT_IDS.scheduleSubmitBtn:
          scheduleAppointment();
          break;
        case ELEMENT_IDS.confirmSubmitBtn:
          sendConfirmation();
          break;
        case ELEMENT_IDS.contractSubmitBtn:
          sendContract();
          break;
        case ELEMENT_IDS.replySubmitBtn:
          sendReply();
          break;
        default:
          break;
      }
    });
  });
}

/**
 * Prepares a reply email.
 * @async
 */
async function sendReply() {
  await createEmail("reply");
}

/**
 * Prepares a contract email.
 * @async
 */
async function sendContract() {
  await createEmail("contract");
}

  /**
   * Prepares a confirmation email.
   * @async
   */
async function sendConfirmation() {
  const type = formState.location.toLowerCase();
  await createEmail(type);
}

/**
 * Schedules an appointment with a lawyer
 * and sends a confirmation email to the client.
 * (TODO: link to confirmation email)
 * @async
 */
async function scheduleAppointment() {
  try {
    // Show loading spinner
    showLoading(true);

    // 1: Validate inputs
    if (!isValidInputs()) {
      throw new Error("Invalid inputs.");
    };

    // 2. Get lawyer
    const lawyer = getLawyer(formState.lawyerId);

    // 3. Fetch calendar events
    const now = new Date();
    const timeRange = {
      // need to be in ISO 8601 format for Microsoft Graph
      startDateTime: now.toISOString(),
      endDateTime: new Date(now.getTime() + 14 * 86400000).toISOString()
    };

    const allEvents = await fetchCalendarEvents(lawyer, timeRange);
    
    if (!allEvents) {
      throw new Error("Failed to fetch calendar events.");
    }

    console.log("Fetched events:", allEvents);

    // 4. Generate candidate slots with breaks
    const candidateSlots = generateSlots(lawyer, allEvents);

    console.log("Generated slots:", candidateSlots.map(s => ({
      start: s.start,
      end: s.end,
      location: s.location
    })));
    
    // 5. Find first valid slot
    const validSlot = candidateSlots.find((slot) => {
      return isValidSlot(lawyer.id, slot, allEvents);
    });

    if (!validSlot) {
      throw new Error("No available slots in next 2 weeks.");
    }

    console.log("Valid slot selected:", validSlot);

    // 6. Create meeting and email
    await createMeeting(validSlot.start, validSlot.end);

  } catch(error) {
    console.error("Scheduling Error:", error);
    showError(error.message);
    throw error;
  } finally {
    // Hide loading spinner
    setTimeout(() => showLoading(false), 1000); // Safety delay
  }
}