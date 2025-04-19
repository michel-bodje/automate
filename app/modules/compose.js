import {
  formState,
  getLawyer,
  getCaseDetails,
  templates,
  loadTemplate,
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
 * @param {Array<string>} category - The category string array.
 */
function setCategory(category) {
  Office.context.mailbox.item.categories.addAsync(category, (addResult) => {
    if (addResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to set category:", addResult.error.message);
    } else {
      console.log(`Category "${category}" set successfully.`);
    }
  });
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
      return "Confirmation of appointment - Allen Madelin";
    }
  }
}

/**
 * Adds taxes to the given amount.
 * @param {number} amount - The amount to which taxes will be added.
 * @param {Boolean} addFOF - Whether or not to add the file opening fee.
 * @returns {number} The total amount with taxes.
 * @throws {Error} If the amount is not a number.
 */
function addTaxes(amount, addFOF = false) {
  if (isNaN(amount)) {
    console.error("Amount is not a number:", amount);
    throw new Error("Amount is not a valid number.");
  }

  // GST + QST + 100$ file opening fee
  // GST: 5% + QST: 9.975% 
  const fof = 100;
  let total = (amount * (1 + 0.05 + 0.09975));
  if (addFOF) total += fof;

  return total;
}

/**
 * Creates an email draft with the specified type and language.
 * @param {string} type - The type of email (e.g., "office", "teams", "phone", "contract" or "reply").
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

    // multilingual support
    const language = formState.clientLanguage === "Français" ? "fr" : "en";
    const template = templates[language][type];

    if (!template) {
      throw new Error(`No template found for type "${type}" in language "${language}".`);
    }

    const lawyer = getLawyer(formState.lawyerId);
    let body = template;

    const appointmentDateTime = new Date(
      `${formState.appointmentDate}T${formState.appointmentTime}`
    );

    // Only validate date and time for appointment confirmations
    if (type === "office" || type === "teams" || type === "phone") {
      if (!appointmentDateTime) {
        throw new Error("No appointment date and time provided.");
      }
      const dateTime = appointmentDateTime;

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

    // Deposit for contract email
    let depositAmount = formState.depositAmount;
    let totalAmount = addTaxes(formState.depositAmount, true);

    depositAmount = Number(depositAmount).toFixed();
    totalAmount = Number(totalAmount).toFixed(2);

    // Adjusted rates for appointment confirmations
    const isFirstConsultation = formState.isFirstConsultation;

    let rates = isFirstConsultation ? 125 : 350;
    let totalRates = addTaxes(rates)

    rates = Number(rates).toFixed();
    totalRates = Number(totalRates).toFixed(2);

    body = body
      .replace("{{lawyerName}}", lawyer.name)
      .replace("{{depositAmount}}", depositAmount)
      .replace("{{totalAmount}}", totalAmount)
      .replace("{{rates}}", rates)
      .replace("{{totalRates}}", totalRates)
    ;

    const subject = getSubject(language, type);
    
    setSubject(subject);
    setRecipient(clientEmail);
    setBody(body);

  } catch (error) {
    console.error("createEmail:", error);
    throw error; // Rethrow the error for further handling if needed
  }
}

/**
 * Creates a meeting draft with the specified details.
 * @param {{ start: Date, end: Date, location: string }} selectedSlot - The proposed time slot.
 */
export async function createMeeting(selectedSlot) {
  try {
    // Fetch the lawyer's details from lawyers.json
    const lawyer = getLawyer(formState.lawyerId);

    // Construct the case details
    const caseDetails = getCaseDetails();
    
    // Construct the subject and body
    const subject = `${formState.clientName} (ma)`;
    const body = `
      <p>Client:&nbsp;&nbsp;&nbsp;&nbsp;${formState.clientName}<br>

      Phone:&nbsp;&nbsp;&nbsp;${formState.clientPhone}<br>

      Email:&nbsp;&nbsp;&nbsp;&nbsp;${formState.clientEmail}<br>

      Lang:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${formState.clientLanguage}</p>

      ${formState.isRefBarreau ? "<p><u><em>Ref. Barreau</em></u></p>" : ""}

      <p>${formState.isFirstConsultation
        ? '<span style="background-color: yellow;">First Consultation</span>'
        : '<span style="background-color: #d3d3d3;">Follow-up</span>'}
      : ${caseDetails}</p>

      <p><strong>Payment</strong>  ${formState.isPaymentMade ? "✔️" : "❌"}<br>
      
      ${formState.isPaymentMade ? `${formState.paymentMethod} (ma)` : ""}</p>
      
      <p>Notes:<br>
      <span style="font-style: italic">${formState.notes}</span></p>
    `;

    // Set details for the draft meeting
    setCategory([lawyer.name]);
    setSubject(subject);
    setMeetingTimes(selectedSlot.start, selectedSlot.end);
    setAttendees([{ displayName: lawyer.name, emailAddress: lawyer.email }]);
    setLocation(selectedSlot.location.charAt(0).toUpperCase() + location.slice(1));
    setBody(body);

  } catch (error) {
    console.error("createMeeting:", error);
    throw error; // Rethrow the error for further handling if needed
  }
}

/**
 * Creates a contract document in Word using the Office.js API.
 * @returns {Promise<void>}
 */
export async function createContract() {
  // Retrieve user inputs from your taskpane UI
  const clientName = formState.clientName;
  const clientEmail = formState.clientEmail;
  const contractTitle = formState.contractTitle;
  let depositAmount = formState.depositAmount;
  let totalAmount = addTaxes(formState.depositAmount, true);

  depositAmount = Number(depositAmount).toFixed();
  totalAmount = Number(totalAmount).toFixed(2);

  // Basic input validation
  if (!clientName || !clientEmail || !contractTitle || !depositAmount) {
    console.log("One or more inputs are missing.");
    return;
  }

  if (!isValidEmail(clientEmail)) {
    console.log("Invalid email format.");
    return;
  }

  const language = formState.clientLanguage === "Français" ? "fr" : "en";

  try {
    // Load the DOCX template
    const doc = await loadTemplate(language);

    // Define a mapping between the placeholders and the input values
    const placeholders = {
      clientName,
      clientEmail,
      contractTitle,
      depositAmount,
      totalAmount,
      date: new Date().toLocaleDateString(),
    };

    // Render the document with the placeholders replaced
    // Generate the final document as a base64 string
    doc.render(placeholders);
    const base64Template = doc.getZip().generate({ type: "base64" });

    // Insert the generated document into Word
    await Word.run(async (context) => {
      // Create new document from the base64 string
      const newDoc = context.application.createDocument(base64Template);
      await context.sync();
      
      // Open in new window
      newDoc.open();
      await context.sync();
      
      // Find all instances of the email
      const searchResults = context.document.body.search(clientEmail, {
        matchCase: false,
        matchWholeWord: true
      });
      searchResults.load("items");
      await context.sync();

      // It's replacing in the first document, not the new one...
      // Replace the email text with a mailto link
      if (searchResults.items.length > 0) {
        const promises = searchResults.items.map(async (range) => {
          range.load("text");
          return context.sync()
            .then(() => {
              // Only replace if it's the exact email match
              if (range.text === clientEmail) {
                range.insertHtml(
                  `<a href="mailto:${clientEmail}">${clientEmail}</a>`,
                  Word.InsertLocation.replace
                );
              }
            });
          }
        );
        await Promise.all(promises);
        await context.sync();
      }
   });
  } catch (error) {
    console.error("Error generating contract:", error);
    throw error;
  }
}