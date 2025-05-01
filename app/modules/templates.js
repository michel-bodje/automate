import teamsTemplateEn from "../../assets/templates/en/Teams.html";
import teamsTemplateFr from "../../assets/templates/fr/Teams.html";
import officeTemplateEn from "../../assets/templates/en/Office.html";
import officeTemplateFr from "../../assets/templates/fr/Office.html";
import phoneTemplateEn from "../../assets/templates/en/Phone.html";
import phoneTemplateFr from "../../assets/templates/fr/Phone.html";
import contractTemplateEn from "../../assets/templates/en/Contract.html";
import contractTemplateFr from "../../assets/templates/fr/Contract.html";
import replyTemplateEn from "../../assets/templates/en/Reply.html";
import replyTemplateFr from "../../assets/templates/fr/Reply.html";
import feedbackTemplateEn from "../../assets/templates/en/Feedback.html";
import feedbackTemplateFr from "../../assets/templates/fr/Feedback.html";
import suiviTemplateEn from "../../assets/templates/en/Suivi.html";
import suiviTemplateFr from "../../assets/templates/fr/Suivi.html";
import teamSignature from "../../assets/templates/en/Equipe Allen Madelin (admin@amlex.ca).htm";

/**
 * A dictionary of html email templates, indexed by language code (e.g., 'en' or 'fr').
 */
export const htmlTemplates = {
  en: {
    teams: teamsTemplateEn,
    office: officeTemplateEn,
    phone: phoneTemplateEn,
    contract: contractTemplateEn,
    reply: replyTemplateEn,
    feedback: feedbackTemplateEn,
    suivi: suiviTemplateEn,
    signature: teamSignature,
  },
  fr: {
    teams: teamsTemplateFr,
    office: officeTemplateFr,
    phone: phoneTemplateFr,
    contract: contractTemplateFr,
    reply: replyTemplateFr,
    feedback: feedbackTemplateFr,
    suivi: suiviTemplateFr,
  },
};

// DOCX templates paths
const docxPaths = {
  en: {
    contract: "assets/templates/en/Contract.docx",
    receipt: "assets/templates/en/PaymentReceipt.docx",
  },
  fr: {
    contract: "assets/templates/fr/Contract.docx",
    receipt: "assets/templates/fr/PaymentReceipt.docx",
  },
};

/**
 * Loads the requested DOCX template for the specified language and prepares it for manipulation.
 * @param {string} language - The language code (e.g., 'en' or 'fr').
 * @param {string} type - The type of Word doc (e.g., 'contract' or 'receipt').
 * @returns {Promise<ArrayBuffer>} - A promise that resolves to the binary content of the template.
 */
export async function loadDocxTemplate(language, type) {
  const templatePath = docxPaths[language][type];

  const content = await fetch(`${process.env.PUBLIC_PATH || ""}${templatePath}`)
    .then((res) => {
      if (!res.ok) {
        throw new Error(`Failed to fetch template at ${templatePath}: ${res.status}`);
      }
      return res.arrayBuffer();
  });

  // Return the binary content of the template
  return content;
}
