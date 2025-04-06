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
        case ELEMENT_IDS.scheduleMode:
          // appointment mode dropdown change
          const manualDate = document.getElementById(ELEMENT_IDS.manualDate);
          const manualTime = document.getElementById(ELEMENT_IDS.manualTime);
          const manualDateLabel = document.querySelector(`label[for=${ELEMENT_IDS.manualDate}]`);
          const manualTimeLabel = document.querySelector(`label[for=${ELEMENT_IDS.manualTime}]`);

          // Show/hide manual date/time inputs based on selected mode          
          if (event.target.value === "manual") {
            manualDateLabel.classList.remove("hidden");
            manualTimeLabel.classList.remove("hidden");
            manualDate.classList.remove("hidden");
            manualTime.classList.remove("hidden");
            manualDate.required = true;
            manualTime.required = true;
          } else {
            manualDateLabel.classList.add("hidden");
            manualTimeLabel.classList.add("hidden");
            manualDate.classList.add("hidden");
            manualTime.classList.add("hidden");
            manualDate.required = false;
            manualTime.required = false;
          }
        case ELEMENT_IDS.confDate:
        case ELEMENT_IDS.confTime:
        case ELEMENT_IDS.manualDate:
        case ELEMENT_IDS.manualTime:
          // date/time input change
          // Determine the active page (confirmation or schedule)
          const activeDateInput = document.getElementById(ELEMENT_IDS.manualDate) || document.getElementById(ELEMENT_IDS.confDate);
          const activeTimeInput = document.getElementById(ELEMENT_IDS.manualTime) || document.getElementById(ELEMENT_IDS.confTime);

          const dateInput = activeDateInput?.value;
          const timeInput = activeTimeInput?.value;

          if (dateInput && timeInput) {
            const dateTime = new Date(`${dateInput}T${timeInput}`);
            console.log("Selected date/time:", dateTime); // debug
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
                font-family: 'Trebuchet MS', 'Tahoma', 'Arial', sans-serif;
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
 * Finds an available time slot for the lawyer's calendar
 * within the next 14 days.
 * @async
 * @returns {Promise<{start: Date, end: Date}>} - The available time slot.
 */
async function findAutoScheduleSlot() {
  const lawyer = getLawyer(formState.lawyerId);
  const location = formState.location;
  const startDateTime = new Date();
  const endDateTime = new Date(startDateTime.getTime() + (14 * 24 * 60 * 60 * 1000));
  // Two weeks from now

  // Fetch calendar events for the lawyer
  const events = await fetchCalendarEvents(lawyer.id, startDateTime, endDateTime);

  // Generate available slots based on the fetched events
  const slots = generateSlots(events, lawyer, location, startDateTime, endDateTime);

  // Return the first valid slot
  const validSlot = find(slots, (slot) =>
    isValidSlot(
      lawyer.id, { start: slot.start, end: slot.end, location: location }, events
    )
  );

  if (!validSlot) {
    throw new Error("No available slots found in next 2 weeks.");
  }
  console.log("Valid slot selected:", validSlot);
  return validSlot;
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

    if (!isValidInputs()) {
      throw new Error("Invalid inputs.");
    };

    const scheduleMode = document.getElementById(ELEMENT_IDS.scheduleMode)?.value;

    let selectedSlot;

    if (scheduleMode === 'manual') {
      const appointmentDateTime = formState.appointmentDateTime;
      if (!appointmentDateTime) {
      throw new Error("Please provide both date and time for manual scheduling.");
      }
      selectedSlot = {
        start: new Date(appointmentDateTime),
        end: new Date(appointmentDateTime.getTime() + (60 * 60 * 1000)),
        // Default to 1 hour duration
      };
      console.log("Scheduled appointment at:", selectedSlot.start);
    } else {
      // Auto-scheduling mode
      selectedSlot = await findAutoScheduleSlot();
      console.log("Auto-scheduled appointment at:", selectedSlot.start);
    }

    // Draft the calendar event
    await createMeeting(selectedSlot.start, selectedSlot.end);

  } catch(error) {
    console.error("Scheduling Error:", error);
    showError(error.message);
    throw error;
  } finally {
    // Hide loading spinner
    setTimeout(() => showLoading(false), 1000); // Safety delay
  }
}