import {
  formState,
  getLawyer,
  getCaseDetails,
  htmlTemplates,
  loadDocxTemplate,
  isValidEmail,
} from "../index.js";
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";

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
  // Check if categories API is available
  if (Office.context.mailbox.item.categories) {
    removeAllCategories(() => addCategory(category));
  } else {
    console.error("Categories API is not available.");
  }

  function getCategories(callback) {
    Office.context.mailbox.item.categories.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to get categories. Error:", result.error.message);
        callback([]);
      } else {
        callback(result.value || []);
      }
    });
  }

  function removeAllCategories(callback) {
    getCategories((categories) => {
      if (categories.length === 0) {
        callback();
        return;
      }
      const categoriesToRemove = categories.map((category) => category.displayName);
      
      Office.context.mailbox.item.categories.removeAsync(categoriesToRemove, (removeResult) => {
        if (removeResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to remove categories. Error:", removeResult.error.message);
        }
        callback();
      });
    });
  }

  function addCategory(category) {
    Office.context.mailbox.item.categories.addAsync(category, (addResult) => {
      if (addResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Failed to set category:", addResult.error.message);
      }
    });
  }
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
    const template = htmlTemplates[language][type];
    const signature = htmlTemplates["en"]["signature"];

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
        .concat(signature)
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
    const lawyer = getLawyer(formState.lawyerId);
    
    // Capitalize location string
    const location = selectedSlot.location.charAt(0).toUpperCase() + selectedSlot.location.slice(1);

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
    setLocation(location);
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
  const { clientName, clientEmail, contractTitle, depositAmount, clientLanguage } = formState;

  // Basic input validation
  if (!clientName || !clientEmail || !contractTitle || !depositAmount) {
    console.error("One or more inputs are missing.");
    return;
  }

  if (!isValidEmail(clientEmail)) {
    console.error("Invalid email format.");
    return;
  }

  const language = clientLanguage === "Français" ? "fr" : "en";

  try {
    // Load the DOCX template as a binary string
    const templateBinary = await loadDocxTemplate(language, "contract");

    // Initialize PizZip with the template binary
    const zip = new PizZip(templateBinary);

    // Initialize Docxtemplater
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Define placeholders
    const placeholders = {
      clientName,
      clientEmail: "{{clientEmail}}", // Placeholder for email link
      contractTitle,
      depositAmount: Number(depositAmount).toFixed(),
      totalAmount: Number(addTaxes(depositAmount, true)).toFixed(2),
      date: new Date().toLocaleDateString(),
    };

    // Render the document
    doc.render(placeholders);

    // Generate the processed document as a base64 string
    const base64Template = doc.getZip().generate({ type: "base64" });

    await Word.run(async (context) => {
      // Create the new document
      const newDoc = context.application.createDocument(base64Template);
      context.trackedObjects.add(newDoc);
      await context.sync();
    
      // Search for the placeholder in the new document
      const searchResults = newDoc.body.search("{{clientEmail}}");
      context.load(searchResults, "items");
      await context.sync();
    
      if (searchResults.items.length === 0) {
        console.error("Placeholder {{clientEmail}} not found in the document.");
        return;
      }
    
      // Replace the placeholder with a mailto hyperlink using HTML
      const mailtoHtml = `<a href="mailto:${clientEmail}">${clientEmail}</a>`;
      searchResults.items[0].insertHtml(mailtoHtml, Word.InsertLocation.replace);
      await context.sync();
    
      console.log("Replaced {{clientEmail}} with a mailto hyperlink using HTML.");

      newDoc.open();
      context.trackedObjects.remove(newDoc);
    });
  } catch (error) {
    console.error("Error creating contract:", error);
  }
}

export async function createReceipt() {
  try {
    const { clientName, lawyerId, depositAmount, clientLanguage } = formState;

    // Basic input validation
    if (!clientName || !lawyerId || !depositAmount) {
      console.error("One or more inputs are missing.");
      return;
    }

    if (!isValidEmail(clientEmail)) {
      console.error("Invalid email format.");
      return;
    }

    const language = clientLanguage === "Français" ? "fr" : "en";

    // Load the DOCX template as a binary string
    const templateBinary = await loadDocxTemplate(language, "receipt");

    // Initialize PizZip with the template binary
    const zip = new PizZip(templateBinary);

    // Initialize Docxtemplater
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Define placeholders
    const placeholders = {
      user,
      clientName,
      amount: Number(depositAmount).toFixed(),
      paymentMethod,
      reason,
      lawyerName: getLawyer(lawyerId).name,
      date: new Date().toLocaleDateString(),
    };

    // Render the document
    doc.render(placeholders);

    // Generate the processed document as a base64 string
    const base64Template = doc.getZip().generate({ type: "base64" });

    await Word.run(async (context) => {
      // Create the new document
      const newDoc = context.application.createDocument(base64Template);
      context.trackedObjects.add(newDoc);
      await context.sync();
      newDoc.open();
      context.trackedObjects.remove(newDoc);
    });
  } catch (error) {
    console.error("Error creating receipt:", error);
  }
}