import {
  ELEMENT_IDS,
  formState,
  getLawyer,
  getAllLawyers,
  getAvailableLocations,
  caseTypeHandlers,
} from "../index.js";

/**
 * Displays the specified page by its ID and hides all other pages.
 *
 * This function assumes that all pages have the class "page" and that the
 * currently visible page has the additional class "active". It removes the
 * "active" class from all pages and adds it to the page with the given ID.
 *
 * @param {string} pageId - The ID of the page to display. The element with this
 * ID should exist in the DOM and have the class "page".
 */
export function showPage(pageId) {
  // Hide all pages
  const pages = document.querySelectorAll(".page");
  pages.forEach(page => page.classList.remove("active"));

  // Show the selected page
  const selectedPage = document.getElementById(pageId);
  if (selectedPage) {
    selectedPage.classList.add("active");

    // If the "Create Contract" page is shown, unhide all elements
    if (pageId === ELEMENT_IDS.wordContractPage) {
      selectedPage.classList.remove("hidden");
    }
  }
}

/** Utility function to reset the page to its initial state. */
export function resetPage() {
  // Clear all form inputs
  const forms = document.querySelectorAll("form");
  forms.forEach(form => form.reset());
  
  // Clear state
  formState.reset();

  // Hide extra case details fields
  hideExtraFields();

  // Reset payment options
  handlePaymentOptions();

  // Navigate back to main menu
  showPage(ELEMENT_IDS.mainPage);

  // Reset dropdowns
  populateLawyerDropdown();
  populateLanguageDropdown();
  if (Office.context.host === Office.HostType.Word) {
    populateContractTitles();
  }
}

/**
 * Sets up the Outlook menu by hiding the "Create Contract" button.
 * This function is called when the application initializes.
 */
export function setupOutlookMenu() {
  // Hide the "Create Contract" button in Outlook
  const createContractBtn = document.getElementById(ELEMENT_IDS.wordContractMenuBtn);
  if (createContractBtn) {
    createContractBtn.classList.add("hidden");
  }
}

/**
 * Sets up the Word menu by hiding all buttons except for the Word contract and user manual buttons.
 * This function is called when the application initializes.
 */
export function setupWordMenu() {
  // Hide all buttons except for the Word contract and user manual buttons
  const menuButtons = document.querySelectorAll('.menu-btn');
  menuButtons.forEach((button) => {
    if (
      button.id !== ELEMENT_IDS.wordContractMenuBtn &&
      button.id !== ELEMENT_IDS.userManualMenuBtn
    ) {
      button.classList.add("hidden");
    }
  });
}

/**
 * Shows or hides the loading overlay.
 *
 * @param {boolean} visible - Whether or not to display the overlay.
 */
export function showLoading(visible) {
  const overlay = document.getElementById("loading-overlay");
  overlay.classList.toggle("hidden", !visible);
}

/**
 * Displays a modal error message window with an "OK" button.
 * @param {string} errorMessage - The error message to display.
 */
export function showErrorModal(errorMessage) {
  const modal = document.getElementById("error-modal");
  const messageElement = document.getElementById("error-message");
  const okButton = document.getElementById("error-ok-button");

  if (!modal || !messageElement || !okButton) {
    console.error("Error modal elements not found.");
    return;
  }

  // Set the error message
  messageElement.textContent = errorMessage;

  // Show the modal
  modal.classList.remove("hidden");

  // Add event listener to the OK button to close the modal
  okButton.onclick = () => {
    modal.classList.add("hidden");
  };
}


/** Utility function to hide all extra form fields and reset their values. */
function hideExtraFields() {
  // Select all extra fields
  const caseDetailsFields = document.querySelectorAll("div[id$='-details']");

  // Clear their input values
  caseDetailsFields.forEach(field => {
    field.hidden = true;
    const inputs = field.querySelectorAll("input, textarea");
    inputs.forEach(input => input.value = "");
  });

  // Hide the entire section
  const caseDetailsElement = document.getElementById(ELEMENT_IDS.caseDetails);

  if (caseDetailsElement) {
    caseDetailsElement.hidden = true;
  }
}

/** Utility function to clear a dropdown and add a placeholder option.
 * @param {HTMLSelectElement} dropdown - The dropdown element to clear.
 * This function clears all existing options and adds a placeholder option.
 * @param {string} placeholder - The placeholder text for the dropdown.
*/
function emptyDropdown(dropdown, placeholder) {
  // Clear existing options
  dropdown.innerHTML = "";

  // Add placeholder option
  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = placeholder;
  placeholderOption.disabled = true;
  placeholderOption.selected = true;
  dropdown.appendChild(placeholderOption);
}

/** 
 * Utility function to populate a dropdown with options.
 * @param {string} elementId - The ID of the dropdown element.
 * @param {Array} options - The options to populate the dropdown with.
 * @param {string} [placeholder="Select an option"] - The placeholder text for the dropdown.
 */
function populateDropdown(elementId, options, placeholder = "Select an option") {
  // Select the dropdown
  const dropdown = document.getElementById(elementId);

  if (!dropdown) {
    console.error(`Dropdown with ID ${elementId} not found.`);
    return;
  }

  emptyDropdown(dropdown, placeholder);

  // Add options
  options.forEach(option => {
    const opt = document.createElement("option");
    opt.value = option.value;
    opt.textContent = option.label;
    dropdown.appendChild(opt);
  });
}

/** Dynamically loads the client language options. */
export function populateLanguageDropdown() {
  // Language dropdown elements
  const elements = [
    ELEMENT_IDS.scheduleClientLanguage,
    ELEMENT_IDS.confClientLanguage,
    ELEMENT_IDS.contractClientLanguage,
    ELEMENT_IDS.replyClientLanguage,
    ELEMENT_IDS.wordClientLanguage,
  ];

  // Populate dropdowns
  elements.forEach(element => {
    populateDropdown(element, [
      { value: "English", label: "English" },
      { value: "Français", label: "Français" },
    ],
    "Select Language");
  });
}

/** Dynamically loads the known lawyers. */
export function populateLawyerDropdown() {
  // Lawyer dropdown elements
  const elements = [
    ELEMENT_IDS.scheduleLawyerId,
    ELEMENT_IDS.confLawyerId,
    ELEMENT_IDS.contractLawyerId,
    ELEMENT_IDS.replyLawyerId,
  ];

  // Get available lawyers
  const lawyers = getAllLawyers();

  // No lawyers found, clear the dropdown
  if (!lawyers || lawyers.length === 0) {
    elements.forEach(element => {
      populateDropdown(element, [], "Select Lawyer");
    });
    return;
  }

  // Populate the dropdowns
  elements.forEach(element =>
    populateDropdown(
      element,
      lawyers.map((lawyer) => ({ value: lawyer.id, label: `${lawyer.name} (${lawyer.id})` })),
      "Select Lawyer"
    )
  );
}

/** Dynamically loads the location options based on the selected lawyer. */
export function populateLocationDropdown() {
  // Location dropdown elements
  const elements = [
    ELEMENT_IDS.scheduleLocation,
    ELEMENT_IDS.confLocation,
  ];
  
  // Get selected lawyer
  const lawyer = getLawyer(formState.lawyerId);
  
  // No lawyer selected, clear the dropdown
  if (!lawyer) {
    elements.forEach(element => {
      populateDropdown(element, [], "Select Location");
    });
    return;
  }

  // Get available locations for the selected lawyer
  const locations = getAvailableLocations(lawyer.id);

  // Populate the dropdowns with the same string for value and label
  elements.forEach(element => {
    populateDropdown(
      element,
      locations.map((location) => ({ value: location, label: location })),
      "Select Location"
    );
  });
}

/** Dynamically loads the case type options based on the selected lawyer. */
export function populateCaseTypeDropdown() {
  // Get selected lawyer
  const lawyer = getLawyer(formState.lawyerId);

  // No lawyer selected, clear the dropdown
  if (!lawyer) {
    populateDropdown(ELEMENT_IDS.caseType, [], "Select Case Type");
    return;
  }

  /**
   * Get available case types for the selected lawyer.
   * If a handler exists for the case type, use its label.
   * Otherwise, use the case type string as the label.
   */
  const caseTypes = lawyer.specialties
    .map(caseType => {
      const handlerExists = caseTypeHandlers[caseType];
      const label = handlerExists
        ? caseTypeHandlers[caseType].label
        : caseType;
      return { value: caseType, label };
    })
  ;

  // Populate the dropdown with the case types
  populateDropdown(
    ELEMENT_IDS.caseType,
    caseTypes,
    "Select Case Type"
  );
}

/**
 * Dynamically populates the contract title dropdown based on the selected language.
 */
export function populateContractTitles() {
  const language = formState.clientLanguage; // Get the selected language from formState
  const contractTitleDropdown = document.getElementById(ELEMENT_IDS.wordContractTitle);

  if (!contractTitleDropdown) {
    console.error("Contract title dropdown not found.");
    return;
  }

  // Define contract title options for each language
  const titles = {
    English: [
      "Representation in Divorce",
      "Representation in Estate Law",
      "Limited Mandate",
    ],
    Français: [
      "Représentation en divorce",
      "Représentation en droit des successions",
      "Mandat Limité",
    ],
  };

  // Get the appropriate titles based on the selected language
  const options = (titles[language] || []).map(title => ({ value: title, label: title }));

  // Add an "Other" option
  options.push({ value: "other", label: "Other (Specify)" });

  // Populate the dropdown
  populateDropdown(ELEMENT_IDS.wordContractTitle, options, "Select Contract Title");
}

/** Handles the case type dropdown change event. */
export function handleCaseDetails() {
  const caseType = formState.caseType;
  const detailsContainer = document.getElementById(ELEMENT_IDS.caseDetails);

  // Hide all case details fields
  hideExtraFields();

  // Dynamically show the selected case details field using the handler
  if (caseTypeHandlers[caseType]) {
    const handlerId = `${caseType}-details`;
    const detailsElement = document.getElementById(handlerId);
    if (detailsElement) {
      detailsContainer.hidden = false;
      detailsElement.hidden = false;
    }
  } else {
    // If no valid case type is selected, hide the entire case details section
    detailsContainer.hidden = true;
  }
}

/**
 * Toggles the visibility of the payment options dropdown based on the "Payment Made" checkbox.
 */
export function handlePaymentOptions() {
  const paymentOptionsContainer = document.getElementById(ELEMENT_IDS.paymentOptionsContainer);

  if (!paymentOptionsContainer) {
    console.error("Payment options container not found.");
    return;
  }

  if (formState.isPaymentMade) {
    paymentOptionsContainer.hidden = false;
  } else {
    paymentOptionsContainer.hidden = true; 
  }
}