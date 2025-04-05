# User Manual for Allen Madelin Outlook Add-In

<div style="text-align: center;">
  <img src="assets/AM-INC.jpg" alt="Corporate Logo">
</div>

## Overview

The Allen Madelin Outlook Add-In (Automate) is designed to streamline appointment scheduling and email drafting for the law firm. It integrates directly into Outlook, allowing users to perform these tasks efficiently without leaving their email client.

This manual is divided into two sections:
1. **User Workflow**: Instructions for using the add-in to schedule appointments and draft emails.
2. **Developer Guide**: A detailed explanation of the codebase to help future developers maintain and extend the add-in.

---

## User Workflow

### Accessing the Add-In
1. Open Outlook.
2. Open a new **email** or **meeting** draft.
3. Select **Automate** from the ribbon options.

<br>
<div style="text-align: center;">
  <img src="assets/ribbon_location_message_AM.png" alt="Ribbon location" width="600">
</div>
<br>

### Scheduling Appointments
1. From the main menu, click **Schedule Appointment**.

<br>
<div style="text-align: center;">
  <img src="assets/mainmenu_AM.png" alt="Main menu" width="200">
</div>
<br>

2. Fill out the required fields in the form:
   - **Client Name**: Enter the client's full name.
   - **Client Phone**: Provide the client's phone number.
   - **Client Email**: Enter the client's email address.
   - **Preferred Language**: Select the client's preferred language (English or French).
   - **Lawyer ID**: Choose the lawyer handling the case.
   - **Preferred Location**: Select the meeting location (Office, Phone, or Teams).
   - **Type of Case**: Choose the type of case (e.g., Divorce, Estate, Employment).
   - Additional details may be required based on the case type (e.g., spouse name for divorce cases).

<br>
<div style="display: flex; justify-content: center; gap: 20px;">
  <img src="assets/appt-scheduler1.png" alt="Page 1 form (1)" width="300">
  <img src="assets/appt-scheduler2.png" alt="Page 1 form (2)" width="300">
</div>
<br>

3. Check any applicable boxes:
   - **Réf. Barreau**: If the client is a referral from the Barreau.
   - **First Consultation**: If this is the client's first consultation.
   - **Payment Made**: If the payment has already been made.
4. Add any **Notes** if necessary.
5. Click **Schedule** to finalize the appointment.

<br>
<div style="text-align: center;">
  <img src="assets/example_schedule2_AM.png" alt="After clicking 'Schedule'" width="600">
</div>
<br>

The add-in will:
- Validate the inputs.
- Check the calendar for a suitable time.
- Fill out the draft meeting with the chosen date and time.

### Writing Draft Emails
1. From the main menu, select the type of email you want to draft:
   - **Send Confirmation**: Draft a confirmation email for an appointment.
   - **Send Contract**: Draft an email with a service contract.
   - **Send Reply**: Draft a reply to a client inquiry.

2. Fill out the required fields in the form:
   - **Client Email**: Enter the recipient's email address.
   - **Preferred Language**: Select the language for the email.
   - **Lawyer ID**: Choose the lawyer associated with the email.
   - Additional fields may appear depending on the email type (e.g., deposit amount for contracts).

3. Click **Create** to generate the draft email.


The add-in will populate the email body using predefined templates and insert the necessary details.

<br>

*Example confirmation email*
<br>
<div style="text-align: center;">
  <img src="assets/example_email_conf_AM.png" alt="Example email conf" width="600">
</div>
<br>

*Example contract email. Note that the amount + tax is automatically calculated.*
<br>
<div style="text-align: center;">
  <img src="assets/example_email_contract_AM.png" alt="Example email contract" width="600">
</div>
<br>

---

## Developer Guide

### Codebase Overview

The add-in is built using JavaScript and integrates with the Office JavaScript API. The codebase is modular, with each module handling a specific aspect of the add-in's functionality.

The source code is hosted on Github at [https://github.com/michel-bodje/automate.git](https://github.com/michel-bodje/automate.git),
and the add-in manifest points to its Github Pages URL at [https://michel-bodje.github.io/automate](https://michel-bodje.github.io/automate).

#### Key Files and Directories
- **`app/`**: Contains the main application logic.
  - **`taskpane.html`**: The main UI for the add-in.
  - **`taskpane.js`**: Handles user interactions and orchestrates the workflow.
  - **`modules/`**: Contains reusable modules for specific functionalities.
    - **`auth.js`**: Handles authentication with Microsoft services.
    - **`graph.js`**: Interacts with the Microsoft Graph API to fetch calendar events.
    - **`compose.js`**: Manages email drafting.
    - **`lawyer.js`**: Manages lawyer data and operations.
    - **`rules.js`**: Implements business rules for scheduling.
    - **`state.js`**: Manages the form state.
    - **`ui.js`**: Handles UI interactions and updates.
    - **`util.js`**: Contains utility functions for validation and formatting.
    - **`timeUtils.js`**: Provides time-related utilities for scheduling.
  - **`templates/`**: Stores email templates in English and French.
- **`manifest.xml`**: Defines the add-in's metadata and configuration.
- **`webpack.config.js`**: Configures the build process.

#### Workflow Explanation
1. **User Interaction**:
   - The user interacts with the UI in `taskpane.html`.
   - Events are handled in `taskpane.js`, which updates the `formState` object and triggers the appropriate actions.

<br>

*Most dropdowns are dynamically generated. Other fields only become visible conditionally.*

```html
<!-- taskpane.html -->

<!-- previous code... -->

<!-- Scheduler page -->
<!--
Creates a meeting for a client.
The form is where client information is entered.
-->
<div id="schedule-page" class="page">
  <h1>Appointment Scheduler</h1>
    <form>
      <label for="schedule-client-name">Client Name:</label>
      <input type="text" id="schedule-client-name" required>

      <label for="schedule-client-phone">Client Phone:</label>
      <input type="tel" id="schedule-client-phone" required>

      <label for="schedule-client-email">Client Email:</label>
      <input type="email" id="schedule-client-email" required 
      <label for="schedule-client-language">Preferred Language:</label>
      <select id="schedule-client-language" required>
          <!-- Options will be dynamically populated by JavaScript -->
      </select  
      <label for="schedule-lawyer-id">Lawyer ID:</label>
      <select id="schedule-lawyer-id" required>
          <!-- Options will be dynamically populated by JavaScript -->
      </select>

      <label for="schedule-location">Preferred Location:</label>
      <select id="schedule-location" required>
          <!-- Options will be dynamically populated by JavaScript -->
      </select>

      <!-- collapsed for brevity... -->  

      <!-- End of page 1 -->
      <button type="submit" id="schedule-appointment-btn  class="submit-btn">Schedule</button>
      <button type="button" class="back-btn">Back</button>
    </form>
</div>

<!-- further code... -->
```
<br>

2. **Scheduling Appointments**:
   - The `scheduleAppointment` function in `taskpane.js`:
     - Validates inputs using `util.js`.
     - Fetches the lawyer's calendar events via `graph.js`.
     - Generates available time slots using `rules.js` and `timeUtils.js`.
     - Creates a meeting using `compose.js`.

3. **Drafting Emails**:
   - The `createEmail` function in `compose.js`:
     - Retrieves the appropriate email template from `templates.js`.
     - Replaces placeholders with dynamic data from `formState`.
     - Sets the email subject, body, and recipients using the Office JavaScript API.

<br>

*In `compose.js`, the `createEmail` and `createMeeting` functions are the basis of this add-in.*
<br>

```js
// compose.js

/**
 * Creates an email draft with the specified type and language.
 * @param {string} type - The type of email (e.g., "office", "teams", "phone", "contract" or "reply").
 */
export async function createEmail(type) {
  try {
    // ...

    // multilingual support
    const language = formState.clientLanguage === "Français" ? "fr" : "en";
    const template = templates[language][type];

    // ...

    const depositAmount = parseFloat(formState.deposit);

    // amount + tax calculation
    const totalAmount = (depositAmount * (1 + 0.05 + 0.09975) + 100).toFixed(2);

    body = body
      .replace("{{lawyerName}}", lawyer.name)
      .replace("{{depositAmount}}", depositAmount)
      .replace("{{totalAmount}}", totalAmount)
    ;

    const subject = getSubject(language, type);
    
    setSubject(subject);
    setRecipient(clientEmail);
    setBody(body);

  } catch (error) {
    console.error("createEmail:", error);
    throw error;
  }
}
```

4. **Authentication**:
   - The `auth.js` module initializes the MSAL library for authentication.
   - Access tokens are acquired to interact with the Microsoft Graph API.

5. **Calendar Integration**:
   - The `graph.js` module fetches and manages calendar events using the Microsoft Graph API.

6. **Business Rules**:
   - The `rules.js` module enforces rules for scheduling, such as avoiding lunch breaks and respecting daily appointment limits.

#### Extending Functionality
- **Adding a New Email Template**:
  1. Create a new HTML file in the `templates/` directory.
  2. Add the template to `templates.js` under the appropriate language.
  3. Update `compose.js` to handle the new template type.

- **Adding a New Case Type**:
  1. Update `lawyerData.json` to include the new case type for relevant lawyers.
  2. Add a handler for the case type in `util.js`.
  3. Update `taskpane.html` to include any additional fields required for the case type.

- **Modifying Business Rules**:
  - Update the `rules.js` module to implement new rules or modify existing ones.

<br>

*Lawyer representation in JSON*

```json
lawyerData.json

{
  "lawyers": [
    {
      "id": "MM",
      "name": "Marie Madelin",
      "email": "marie.madelin@amlex.ca",
      "workingHours": {"start": "9:00", "end": "17:00"},
      "breakMinutes": 0,
      "maxDailyAppointments": 5,
      "specialties": ["estate", "mandates", "common"]
    },
    {
      "id": "DH",
      "name": "Dorin Holban",
      "email": "dorin.holban@amlex.ca",
      "workingHours": {"start": "10:30", "end": "17:00"},
      "breakMinutes": 30,
      "maxDailyAppointments": 5,
      "specialties": ["divorce", "employment", "business", "common"]
    },
    {
      "id": "TG",
      "name": "Tim Gagin",
      "email": "tim.gagin@amlex.ca",
      "workingHours": {"start": "9:30", "end": "17:00"},
      "breakMinutes": 30,
      "maxDailyAppointments": 5,
      "specialties": ["estate", "real_estate", "defamations", "contract", "common"]
    },

    collapsed for brevity...
  ]
}
```
<br>

*Editing case types*

```js
// util.js

/** Handles the case type details based on the selected case type. */
export const caseTypeHandlers = {
  divorce: {
    label: "Divorce / Family Law",
    handler: function () {
      const spouseName = document.getElementById(ELEMENT_IDS.spouseName).value;
      const conflictSearchDone = document.getElementById(ELEMENT_IDS.conflictSearchDoneDivorce).checked;
      return `
        ${this.label}
        <p><strong>Spouse Name:</strong> ${spouseName}</p>
        <p>Conflict Search Done? ${conflictSearchDone ? "✔️" : "❌"}</p>
      `;
    },
  },
  estate: {
    label: "Successions / Estate Law",
    handler: function () {
      const deceasedName = document.getElementById(ELEMENT_IDS.deceasedName).value;
      const executorName = document.getElementById(ELEMENT_IDS.executorName).value;
      const conflictSearchDone = document.getElementById(ELEMENT_IDS.conflictSearchDoneEstate).checked;
      return `
        ${this.label}
        <p><strong>Deceased Name:</strong> ${deceasedName}</p>
        <p><strong>Executor Name:</strong> ${executorName}</p>
        <p>Conflict Search Done? ${conflictSearchDone ? "✔️" : "❌"}</p>
      `;
    },
  },

  // collapsed for brevity...
}
```
<br>

#### Debugging and Testing
- To run the add-in locally, use the following npm scripts:
    1. Start the development server: `npm start`
    2. Stop the development server: Press `Ctrl + C` in the terminal where the server is running.
- Ensure that all dependencies are installed by running `npm install` before starting the server.
- Use `npm run lint` to check for code quality issues.
- Test the add-in in both development and production environments to ensure compatibility.  
The `webpack.config.js` file is already configured to handle differences between development and production environments. To test the production build:
    1. Run `npm run build` to generate the production files.
    2. Access the production version of the add-in via Outlook logged to [admin@amlex.ca](mailto:admin@amlex.ca).

<br>
<div style="text-align: center;">
  <img src="assets/webpack.png" alt="Web server" width="600">
</div>
<br>

---

## Conclusion

This add-in simplifies appointment scheduling and email drafting for Allen Madelin. The modular codebase ensures maintainability and extensibility, allowing future developers to adapt the add-in to evolving business needs.