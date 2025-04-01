import teamsTemplateEn from "../templates/en/Teams.html";
import teamsTemplateFr from "../templates/fr/Teams.html";
import officeTemplateEn from "../templates/en/Office.html";
import officeTemplateFr from "../templates/fr/Office.html";
import phoneTemplateEn from "../templates/en/Phone.html";
import phoneTemplateFr from "../templates/fr/Phone.html";
import contractTemplateEn from "../templates/en/Contract.html";
import contractTemplateFr from "../templates/fr/Contract.html";
import replyTemplateEn from "../templates/en/Reply.html";
import replyTemplateFr from "../templates/fr/Reply.html";
import feedbackTemplateEn from "../templates/en/Feedback.html";
import feedbackTemplateFr from "../templates/fr/Feedback.html";
import suiviTemplateEn from "../templates/en/Suivi.html";
import suiviTemplateFr from "../templates/fr/Suivi.html";

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