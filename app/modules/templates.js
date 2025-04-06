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
 * A dictionary of email templates, indexed by language and email type.
 * @type {Object.<string, Object.<string, string>>}
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