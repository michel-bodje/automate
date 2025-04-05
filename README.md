# Automate for Allen Madelin

An Outlook Office Add-in for law firms.

## Setup locally (development)
1. Clone this repo.
2. Run `npm install`.
3. Run `npx office-addin-dev-settings m365-account login` and select your Office work or school account.
3. Run `npm start` to test on your desktop.

## Setup company-wide (production)
This assumes that all targeted computers have access to the same Office account.

1. Verify manifest.xml file contains a link to "https://michel-bodje.github.io/automate/".
2. In the add-in menu (https://outlook.office.com/mail/inclientstore), upload the manifest.xml as a new custom add-in.
3. Enjoy!
