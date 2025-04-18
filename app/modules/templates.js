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
import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";

/**
 * A dictionary of templates, indexed by language and type.
 */
export const templates = {
  en: {
    teams: teamsTemplateEn,
    office: officeTemplateEn,
    phone: phoneTemplateEn,
    contract: contractTemplateEn,
    reply: replyTemplateEn,
    feedback: feedbackTemplateEn,
    suivi: suiviTemplateEn,
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

// The path to the DOCX contract templates
const docxContractPath = {
  en: "assets/templates/en/Contract.docx",
  fr: "assets/templates/fr/Contract.docx",
};

/**
 * Loads the DOCX contract template for the specified language and prepares it for manipulation.
 * @param {string} language - The language code (e.g., 'en' or 'fr').
 */
export async function loadTemplate(language) {
  const templatePath = docxContractPath[language];

  const content = await fetch(`${process.env.PUBLIC_PATH || ""}${templatePath}`)
    .then(res => {
      if (!res.ok) {
      throw new Error(`Failed to fetch template at ${templatePath}: ${res.status}`);
  }
  return res.arrayBuffer();
});

  // Load the DOCX file into PizZip
  const zip = new PizZip(content)

  // Initialize Docxtemplater with the zip content
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  return doc; // Return the Docxtemplater instance for further use
}
