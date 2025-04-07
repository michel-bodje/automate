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
    docxContract: "",
  },
  fr: {
    teams: teamsTemplateFr,
    office: officeTemplateFr,
    phone: phoneTemplateFr,
    contract: contractTemplateFr,
    reply: replyTemplateFr,
    feedback: feedbackTemplateFr,
    suivi: suiviTemplateFr,
    docxContract: "",
  },
};

// The path to the DOCX contract templates
// TODO: Test these paths in production
const docxContractPath = {
  en: "../../assets/templates/en/Contract.docx",
  fr: "../../assets/templates/fr/Contract.docx",
};

/**
 * Loads the DOCX contract template for the specified language and converts it to a base64 string.
 * This function is called when the application initializes.
 * @param {string} language - The language code (e.g., 'en' or 'fr').
 */
export async function loadTemplate(language) {
  const templatePath = docxContractPath[language] || null;
  if (!templatePath) {
    throw new Error(`Template not found for language: ${language}`);
  }
  
  try {
    const base64Template = await getTemplateBase64(templatePath);
    // Store the base64 string in the templates object
    templates[language].docxContract = base64Template;
  } catch (error) {
    console.error('Failed to load template:', error);
  }
}

/**
 * Fetches a template file from the specified path and converts it to a base64 string.
 * @param {string} templatePath - The path to the template file.
 * @returns {Promise<string>} A promise that resolves to the base64 string of the template.
 */
async function getTemplateBase64(templatePath) {
  try {
    const response = await fetch(templatePath);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const blob = await response.blob();
    return await convertBlobToBase64(blob);
  } catch (error) {
    console.error('Error fetching and converting the template:', error);
    throw error;
  }
}

/**
 * Converts a Blob object to a base64 string.
 * @param {Blob} blob - The Blob object to convert.
 * @returns {Promise<string>} A promise that resolves to the base64 string.
 */
function convertBlobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      // reader.result is a data URL of the form "data:<mime-type>;base64,<data>"
      // We split it to get only the base64 string.
      const base64data = reader.result.split(',')[1];
      resolve(base64data);
    };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}