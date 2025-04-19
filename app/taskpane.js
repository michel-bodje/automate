import "./styles/main.css";
import {
  ELEMENT_IDS,
  formState,
  getLawyer,
  showPage,
  resetPage,
  initTaskpaneWord,
  initTaskpaneOutlook,
  populateLawyerDropdown,
  populateContractTitles,
  handleCaseDetails,
  handlePaymentOptions,
  isValidInputs,
  isValidSlot,
  isSameDay,
  generateSlots,
  msalInstance,
  fetchCalendarEvents,
  createEmail,
  createMeeting,
  createContract,
  showLoading,
  showErrorModal,
  openPopup,
} from "./index.js";

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    // Initialize MSAL for authentication
    await msalInstance.initialize();
    // Setup taskpane UI for Outlook
    initTaskpaneOutlook();
  } else if (info.host === Office.HostType.Word) {
    // Setup taskpane UI for Word
    initTaskpaneWord();
  }
  attachEventListeners();
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
          break;

        case ELEMENT_IDS.scheduleLocation:
        case ELEMENT_IDS.confLocation:
          // location dropdown change
          formState.update("location", value);
          break;

        case ELEMENT_IDS.caseType:
          // case type dropdown change
          formState.update("caseType", value);
          populateLawyerDropdown();
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

        case ELEMENT_IDS.scheduleFirstConsultation:
        case ELEMENT_IDS.confFirstConsultation:
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
          // Open user manual in new tab
          openPopup({
            title: "User Manual",
            contentOrFile: "./user-manual.html",
            isFile: true,
            position: "bottom-right",
          });
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
 * Displays a popup with available calendar slots for the next 2 weeks,
 * with the selected slot at the top. The popup is only displayed when
 * auto-scheduling is enabled.
 * @param {Array<{start: Date, end: Date}>} validSlots - The array of valid slot objects with start and end times.
 * @param {{start: Date, end: Date}} selectedSlot - The selected slot object with start and end times.
 */
function popupAvailableSlots(validSlots, selectedSlot) {
  let popupContent = "<h3>Available calendar slots for the next 2 weeks</h3><ul>";

  // Ensure the selected slot is displayed first
  const selectedSlotIndex = validSlots.indexOf(selectedSlot);

  if (selectedSlotIndex !== -1) {
    const slot = validSlots[selectedSlotIndex];
    popupContent += formatSlot(slot, true);
  }

  // Display all future slots after the selected slot
  let previousDate = selectedSlot.start.toLocaleDateString();
  for (let i = selectedSlotIndex + 1; i < validSlots.length; i++) {
    const slot = validSlots[i];

    // Insert a line break if the day changes
    if (slot.start.toLocaleDateString() !== previousDate) {
      popupContent += "<br>";
      previousDate = slot.start.toLocaleDateString();
    }

    popupContent += formatSlot(slot);
  }

  popupContent += "</ul>";

  openPopup({
    title: "Available Slots",
    contentOrFile: popupContent,
    width: 480,
    height: 300,
    position: "bottom-right",
  });

  /**
   * Formats a slot into an HTML list item.
   * @param {Object} slot - The slot object with start and end times.
   * @param {boolean} [isSelected=false] - Whether the slot is the selected one.
   * @returns {string} - The formatted HTML string for the slot.
  */
  function formatSlot(slot, isSelected = false) {
    const startDate = slot.start.toLocaleString([], { weekday: "long", year: "numeric", month: "long", day: "numeric" });
    const startTime = slot.start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    const endTime = slot.end.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

    return `<li>${isSelected ? "<strong>Selected Slot:</strong><br>" : ""}${startDate} - ${startTime} to ${endTime}</li>${isSelected ? "<br>" : ""}`;
  }
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
  const type = formState.location;
  await createEmail(type);
}

/**
 * Returns valid time slots for the appointment within the accepted range.
 * @param {string} scheduleMode - The schedule mode ("auto" or "manual").
 * @async
 * @returns {Promise<Array<{ start: Date, end: Date, location:string }>>} An array of valid time slots.
 */
async function findSlots(scheduleMode) {
  try {
    const lawyer = getLawyer(formState.lawyerId);
    const location = formState.location;
    const events = await fetchCalendarEvents();

    if (scheduleMode === "manual") {
      // Retrieve manually selected date and time
      const manualDate = document.getElementById(ELEMENT_IDS.manualDate)?.value;
      const manualTime = document.getElementById(ELEMENT_IDS.manualTime)?.value;

      if (!manualDate || !manualTime) {
        throw new Error("Please select a valid date and time for manual scheduling.");
      }

      const start = new Date(`${manualDate}T${manualTime}`);
      const end = new Date(start.getTime() + 60 * 60 * 1000); // Assume 1-hour duration

      const selectedSlot = { start, end, location };

      // Validate the manually selected slot
      if (!isValidSlot(lawyer.id, selectedSlot, events)) {
        throw new Error("The selected time slot is not valid.");
      }

      return [selectedSlot]; // Return the valid manual slot as an array
    }

    // Auto-scheduling: Generate and validate slots
    const slots = generateSlots(lawyer, location, events);
    const validSlots = slots.filter(slot =>
      isValidSlot(lawyer.id, slot, events)
    );

    if (validSlots.length === 0) {
      throw new Error("No available slots found in the next 2 weeks.");
    }

    console.log("Valid slots:", validSlots);
    return validSlots;
  } catch (error) {
    console.error("Error finding schedule slot:", error);
    showErrorModal(error.message);
  }
}

/**
 * Schedules an appointment based on the selected inputs.
 * @async
 * @throws {Error} if the inputs are invalid or if no available slots are found.
 */
async function scheduleAppointment() {
  try {
    // Show loading spinner
    showLoading(true);

    if (!isValidInputs()) {
      throw new Error("Invalid inputs.");
    }

    const scheduleMode = document.getElementById(ELEMENT_IDS.scheduleMode)?.value;

    // Fetch valid slots based on the schedule mode
    const validSlots = await findSlots(scheduleMode);

    if (!validSlots || validSlots.length === 0) {
      throw new Error("No available slots found.");
    }

    // Select the appropriate slot
    const now = new Date();
    const selectedSlot =
      scheduleMode === "manual"
        ? validSlots[0] // Manual mode returns a single valid slot
        : validSlots.find(slot => !isSameDay(slot.start, now) || slot.start > now) || validSlots[0]; // Auto mode logic

    console.log("Selected appointment slot:", selectedSlot.start);

    // Display available slots in a popup for auto-scheduling
    if (scheduleMode === "auto") {
      popupAvailableSlots(validSlots, selectedSlot);
    }

    // Create the appointment in the lawyer's calendar
    await createMeeting(selectedSlot);
    console.log("Scheduled appointment successfully.");
  } catch (error) {
    console.error("Scheduling Error:", error);
    showErrorModal(error.message);
  } finally {
    // Hide loading spinner
    showLoading(false);
  }
}