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
  }
}

/** Utility function to reset the page to its initial state. */
export function resetPage() {
  // Clear all form inputs
  const forms = document.querySelectorAll("form");
  forms.forEach(form => form.reset());
  
  // Clear state
  formState.reset();

  // Reset dropdowns to their placeholder values
  const dropdowns = document.querySelectorAll("select");
  dropdowns.forEach(dropdown => {
    dropdown.selectedIndex = 0;
  });

  // Hide extra case details fields
  hideExtraFields();

  // Depopulate lawyer-specific dropdowns in-between resets
  populateLocationDropdown();
  populateCaseTypeDropdown();

  // Reset payment options
  handlePaymentOptions();

  // Navigate back to main menu
  showPage(ELEMENT_IDS.mainPage);
}

/** Utility function to hide all extra form fields and reset their values. */
export function hideExtraFields() {
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
 * Shows an error message in the UI for a short duration.
 *
 * @param {string} message - The error message to display.
 */
export function showError(message) {
  const errorBar = document.getElementById("error-message");
  errorBar.querySelector(".ms-MessageBar-text").textContent = message;
  errorBar.classList.remove("hidden");
  setTimeout(() => errorBar.classList.add("hidden"), 5000);
}

/** 
 * Utility function to populate a dropdown with options.
 * @param {string} elementId - The ID of the dropdown element.
 * @param {Array} options - The options to populate the dropdown with.
 * @param {string} [placeholder="Select an option"] - The placeholder text for the dropdown.
 */
export function populateDropdown(elementId, options, placeholder = "Select an option") {
  // Select the dropdown
  const dropdown = document.getElementById(elementId);

  if (!dropdown) {
    console.error(`Dropdown with ID ${elementId} not found.`);
    return;
  }

  // Clear existing options
  dropdown.innerHTML = "";

  // Add placeholder option
  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = placeholder;
  placeholderOption.disabled = true;
  placeholderOption.selected = true;
  dropdown.appendChild(placeholderOption);

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