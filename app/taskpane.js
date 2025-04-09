import "./styles/main.css";
import {
  ELEMENT_IDS,
  formState,
  getLawyer,
  showPage,
  resetPage,
  setupWordMenu,
  setupOutlookMenu,
  populateLawyerDropdown,
  populateLanguageDropdown,
  populateLocationDropdown,
  populateCaseTypeDropdown,
  populateContractTitles,
  handleCaseDetails,
  handlePaymentOptions,
  isValidInputs,
  isValidSlot,
  generateSlots,
  msalInstance,
  fetchCalendarEvents,
  createEmail,
  createMeeting,
  createContract,
  showLoading,
  showErrorModal,
} from "./index.js";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize MSAL for authentication
    await msalInstance.initialize();
    // Setup UI components for Outlook
    setupOutlookMenu();
    populateLawyerDropdown();
    populateLanguageDropdown();
    attachEventListeners();
  } else if (info.host === Office.HostType.Word) {
    // Setup UI components for Word
    setupWordMenu();
    populateLanguageDropdown();
    populateContractTitles();
    attachEventListeners();
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
        case ELEMENT_IDS.wordClientName:
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
        case ELEMENT_IDS.wordClientEmail:
          // client email input change
          formState.update("clientEmail", value);
          break;
        case ELEMENT_IDS.scheduleClientLanguage:
        case ELEMENT_IDS.confClientLanguage:
        case ELEMENT_IDS.contractClientLanguage:
        case ELEMENT_IDS.replyClientLanguage:
        case ELEMENT_IDS.wordClientLanguage:
          // client language dropdown change
          formState.update("clientLanguage", value);
          if (Office.context.host === Office.HostType.Word) {
            populateContractTitles();
          }
          break;
        case ELEMENT_IDS.scheduleMode:
          // appointment mode dropdown change
          const manualDate = document.getElementById(ELEMENT_IDS.manualDate);
          const manualTime = document.getElementById(ELEMENT_IDS.manualTime);
          const manualDateLabel = document.querySelector(`label[for=${ELEMENT_IDS.manualDate}]`);
          const manualTimeLabel = document.querySelector(`label[for=${ELEMENT_IDS.manualTime}]`);

          // Show/hide manual date/time inputs based on selected mode          
          if (event.target.value === "auto") {
            manualDateLabel.classList.add("hidden");
            manualTimeLabel.classList.add("hidden");
            manualDate.classList.add("hidden");
            manualTime.classList.add("hidden");
            manualDate.required = false;
            manualTime.required = false;
          } else {
            manualDateLabel.classList.remove("hidden");
            manualTimeLabel.classList.remove("hidden");
            manualDate.classList.remove("hidden");
            manualTime.classList.remove("hidden");
            manualDate.required = true;
            manualTime.required = true;
          }
          break;
        case ELEMENT_IDS.confDate:
        case ELEMENT_IDS.manualDate:
          console.log("Date changed:", value);
          formState.update("appointmentDate", value);
          break;
        case ELEMENT_IDS.confTime:
        case ELEMENT_IDS.manualTime:
          console.log("Time changed:", value);
          formState.update("appointmentTime", value); 
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
        case ELEMENT_IDS.emailContractDeposit:
        case ELEMENT_IDS.wordContractDeposit:
          // contract deposit input change
          formState.update("depositAmount", value);
          break;
        case ELEMENT_IDS.wordContractTitle:
          // Show/hide custom contract title input based on selection
          const customTitleInput = document.getElementById(ELEMENT_IDS.customContractTitle);
          if (value === "other") {
            customTitleInput.classList.remove("hidden");
            customTitleInput.required = true;
          } else {
            customTitleInput.classList.add("hidden");
            customTitleInput.required = false;
            formState.update("contractTitle", value);
          }
          break;
        case ELEMENT_IDS.customContractTitle:
          // Update form state with custom contract title
          formState.update("contractTitle", value);
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
        case ELEMENT_IDS.wordContractMenuBtn:
          showPage(ELEMENT_IDS.wordContractPage);
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
        case ELEMENT_IDS.wordContractSubmitBtn:
          createContract();
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
 * Finds all available time slots for the lawyer's calendar
 * within the next 14 days.
 * @async
 * @returns {Promise<Array<{start: Date, end: Date}>>} - An array of valid time slots.
 */
async function findAutoScheduleSlots() {
  try {
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
    // TODO: add a more sophisticated slot selection algorithm
    // e.g., based on client preferences, lawyer availability, etc.
    const validSlots = slots.filter((slot) =>
      isValidSlot(
        lawyer.id, { start: slot.start, end: slot.end, location: location }, events
      )
    );
  
    if (validSlots.length === 0) {
      throw new Error("No available slots found in the next 2 weeks.");
    }
    console.log("Valid slots:", validSlots);
    return validSlots;
} catch (error) {
    console.error("Error finding auto-schedule slot:", error);
    throw error;
  }
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
      const appointmentDateTime = new Date(
        `${formState.appointmentDate}T${formState.appointmentTime}`
      );
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
      const validSlots = await findAutoScheduleSlots();
      
      // Use the first valid slot
      selectedSlot = validSlots[0]; 

      // TODO: Show the valid slots to the user for selection here
      // (e.g., in a dropdown or modal)

      // Display the slots in a popup
      const popupWindow = window.open("", "Available Slots", "width=400,height=300");
      popupWindow.document.write("<html><head><title>Available Slots</title></head><body>");
      popupWindow.document.write("<h3>Available Slots</h3>");
      popupWindow.document.write("<ul>");
      validSlots.forEach(slot => {
        const startDate = slot.start.toLocaleDateString();
        const startTime = slot.start.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const endTime = slot.end.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        popupWindow.document.write(`<li>Date: ${startDate} - Time: ${startTime} to ${endTime}</li>`);
      });
      popupWindow.document.write("</ul>");
      popupWindow.document.write("<button onclick='window.close()'>Close</button>");
      popupWindow.document.write("</body></html>");
      popupWindow.document.close();

      // Apply the same style as in the user manual
      const style = popupWindow.document.createElement('style');
      style.textContent = `
        body {
          background-color: #1e1e1e;
          color: #f4f4f4;
          font-family: 'Trebuchet MS', 'Tahoma', 'Arial', sans-serif;
          line-height: 1.6;
          margin: 0;
          padding: 1rem;
        }
        h3 {
          color: #0078d4;
        }
        ul {
          list-style-type: none;
          padding: 0;
        }
        li {
          margin: 0.5rem 0;
        }
        button {
          background-color: #0078d4;
          color: #fff;
          border: none;
          padding: 0.5rem 1rem;
          font-size: 1rem;
          cursor: pointer;
          border-radius: 4px;
        }
        button:hover {
          background-color: #005a9e;
        }
      `;
      popupWindow.document.head.appendChild(style);
      
      console.log("Auto-scheduled appointment at:", selectedSlot.start);
    }

    // Create the appointment in the lawyer's calendar
    await createMeeting(selectedSlot.start, selectedSlot.end);
    console.log("Scheduled appointment successfully.");
  } catch(error) {
    console.error("Scheduling Error:", error.message);
    showErrorModal(error.message);
  } finally {
    // Hide loading spinner
    showLoading(false);
  }
}