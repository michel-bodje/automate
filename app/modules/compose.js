import {
  formState,
  getLawyer,
  getCaseDetails,
  templates,
  isValidEmail,
} from "../index.js";

/**
 * Sets the subject in the draft email.
 * @param {string} subject - The email subject.
 */
function setSubject(subject) {
  Office.context.mailbox.item.subject.setAsync(subject, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set subject:", asyncResult.error.message);
    }
  });
}

/**
 * Sets the body in the draft email.
 * @param {string} body - The email body.
 */
function setBody(body) {
  Office.context.mailbox.item.body.setAsync(
    body,
    { coercionType: Office.CoercionType.Html },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set body:", asyncResult.error.message);
      }
    }
  );
}

/**
 * Sets the recipient in the draft email.
 * @param {string} recipient - The recipient's email address.
 */
function setRecipient(recipient) {
  Office.context.mailbox.item.to.setAsync([recipient], (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set recipient:", asyncResult.error.message);
    }
  });
}

/**
 * Sets the attendees in the draft meeting.
 * @param {Array} attendees - An array of attendee objects with displayName and emailAddress.
 */
function setAttendees(attendees) {
  Office.context.mailbox.item.requiredAttendees.setAsync(attendees, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set attendees:", asyncResult.error.message);
    }
  });
}

/**
 * Sets the category in the draft meeting.
 * @param {string} category - The category string array.
 *
 */
function setCategory(category) {
  // Function to remove all existing categories
  function removeAllCategories(callback) {
    Office.context.mailbox.item.categories.getAsync((getResult) => {
      if (getResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to retrieve existing categories:", getResult.error.message);
        callback(); // Proceed to add the new category even if retrieval fails
      } else {
        const currentCategories = getResult.value;

        if (currentCategories && currentCategories.length > 0) {
          let pendingRemovals = currentCategories.length;

          currentCategories.forEach((cat) => {
            Office.context.mailbox.item.categories.removeAsync(cat, (removeResult) => {
              if (removeResult.status === Office.AsyncResultStatus.Failed) {
                console.error(`Failed to remove category "${cat}":`, removeResult.error.message);
              }

              pendingRemovals--;
              if (pendingRemovals === 0) {
                callback(); // All categories removed, proceed to add the new category
              }
            });
          });
        } else {
          callback(); // No categories to remove, proceed to add the new category
        }
      }
    });
  }

  // Function to add the new category
  function addCategory() {
    Office.context.mailbox.item.categories.addAsync(category, (addResult) => {
      if (addResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set category:", addResult.error.message);
      } else {
        console.log(`Category "${category}" set successfully.`);
      }
    });
  }

  // First remove all categories, then add the new one
  removeAllCategories(addCategory);
}

/**
 * Sets the location in the draft meeting.
 * @param {string} location - The location to set for the meeting.
 */
function setLocation(location) {
  Office.context.mailbox.item.location.setAsync(location, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set location:", asyncResult.error.message);
    } else {
      console.log("Location set successfully:", location);
    }
  });
}

/**
 * Sets the start and end times for the draft meeting.
 * @param {Date} startTime - The start time of the meeting.
 * @param {Date} endTime - The end time of the meeting.
 */
function setMeetingTimes(startTime, endTime) {
  Office.context.mailbox.item.start.setAsync(startTime, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set start time:", asyncResult.error.message);
    }
  });

  Office.context.mailbox.item.end.setAsync(endTime, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set end time:", asyncResult.error.message);
    }
  });
}

/**
 * Returns the email subject line for the given language and type.
 * @param {string} language - The language of the email.
 * @param {string} type - The type of email.
 * @returns {string} The subject line.
 */
function getSubject(language, type) {
  if (language === "fr") {
    if (type === "contract") {
      return "Contrat de services - Allen Madelin"
    } else if (type === "reply") {
      return "Réponse - Allen Madelin";
    } else {
      return "Confirmation de rendez-vous - Allen Madelin";
    }
  } else {
    if (type === "contract") {
      return "Contract of services - Allen Madelin";
    } else if (type === "reply") {
      return "Reply - Allen Madelin";
    } else {
      return "Appointment Confirmation - Allen Madelin";
    }
  }
}

/**
 * Creates an email draft with the specified type and language.
 * @param {string} type - The meeting type of email (e.g., "office", "teams", "phone").
 */
export async function createEmail(type) {
  try {
    const clientEmail = formState.clientEmail;
    if (!clientEmail) {
      throw new Error("No client email provided.");
    }
    if (!isValidEmail(clientEmail)) {
      throw new Error("Please provide a valid email address.");
    }

    const language = formState.clientLanguage === "Français" ? "fr" : "en";
    const template = templates[language][type];
    if (!template) {
      throw new Error(`No template found for type "${type}" in language "${language}".`);
    }

    const lawyer = getLawyer(formState.lawyerId);
    let body = template;

    // Only validate date and time for appointment confirmations
    if (type === "office" || type === "teams" || type === "phone") {
      if (!formState.appointmentDateTime) {
        throw new Error("No appointment date and time provided.");
      }
      const dateTime = formState.appointmentDateTime;

      const date = dateTime.toLocaleDateString(language == "fr" ? "fr-CA" : "en-US", {
        weekday: "long",
        day: "numeric",
        month: "long",
        year: "numeric",
      });
  
      const time = dateTime.toLocaleTimeString(language == "fr" ? "fr-CA" : "en-US", {
        hour: "2-digit",
        minute: "2-digit",
      });

      body = body
        .replace("{{date}}", date)
        .replace("{{time}}", time)
      ;
  
    }

    const depositAmount = parseFloat(formState.deposit);
    const totalAmount = (depositAmount * (1 + 0.05 + 0.09975) + 100).toFixed(2);

    body = body
      .replace("{{lawyerName}}", lawyer.name)
      .replace("{{depositAmount}}", depositAmount)
      .replace("{{totalAmount}}", totalAmount);

    const subject = getSubject(language, type);
    
    setSubject(subject);
    setRecipient(clientEmail);
    setBody(body);

  } catch (error) {
    console.error("createEmail:", error);
  }
}

/**
 * Creates a meeting draft with the specified details.
 * @param {Date} startTime - The start time of the appointment.
 * @param {Date} endTime - The end time of the appointment.
 */
export async function createMeeting(startTime, endTime) {
  try {
    // Fetch the lawyer's details from lawyers.json
    const lawyer = getLawyer(formState.lawyerId);

    // Construct the case details
    const caseDetails = getCaseDetails();
    
    // Construct the subject and body
    const subject = `${formState.clientName} (ma)`;
    const body = `
      <p><strong>Client:</strong>    ${formState.clientName}<br>

      <strong>Phone:</strong>    ${formState.clientPhone}<br>

      <strong>Email:</strong>    ${formState.clientEmail}<br>

      <strong>Lang:</strong>    ${formState.clientLanguage}</p>

      ${formState.isRefBarreau ? "<p><u><em>Ref. Barreau</em></u></p>" : ""}

      <p>${formState.isFirstConsultation ? '<span style="background-color: yellow;">First Consultation</span>' : "Follow-up"}</p>
      
      <p>${caseDetails}</p>

      <p><strong>Payment</strong>  ${formState.isPaymentMade ? "✔️" : "❌"}<br>
      
      ${formState.isPaymentMade ? `${formState.paymentMethod} (ma)` : ""}</p>
      
      <p>Notes:<br>
      <span style="font-style: italic">${formState.notes}</span></p>
    `;

    // Set details for the draft meeting
    setSubject(subject);
    setBody(body);
    setLocation(formState.location);
    setAttendees([{ displayName: lawyer.name, emailAddress: lawyer.email }]);
    setCategory([lawyer.name]);

    // The core function: set meeting date and time
    setMeetingTimes(startTime, endTime);
  } catch (error) {
    console.error("createMeeting:", error.message);
  }
}