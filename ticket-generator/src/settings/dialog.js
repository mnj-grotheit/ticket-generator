/* eslint-disable no-undef */
/* global Office, console */

import { callAiService } from "../helper/aiService.js";

// Hilfsfunktion: Gibt den E-Mail-Body als Promise zurück
function getEmailBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

// Vorbefüllen der Formularfelder mit den E-Mail-Daten
async function prefillForm() {
  try {
    const item = Office.context.mailbox.item;
    const subject = item.subject || "";
    const emailBody = await getEmailBodyAsync();
    const sender = item.from && item.from.emailAddress ? item.from.emailAddress : "Unbekannt";
    const recipients = item.to || [];
    const recipient = recipients.length > 0 ? recipients[0].emailAddress : "Unbekannt";

    document.getElementById("titel").value = subject;
    document.getElementById("beschreibung").value = emailBody;
    document.getElementById("verantwortlicher").value = recipient;
    document.getElementById("ansprechpartner").value = sender;
  } catch (error) {
    console.error("Fehler beim Vorbefüllen des Formulars:", error);
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Outlook Add-In ist bereit.");
    prefillForm();
    document.getElementById("ticketForm").addEventListener("submit", handleFormSubmit);
  }
});

async function handleFormSubmit(event) {
  event.preventDefault();
  const statusMessage = document.getElementById("statusMessage");
  statusMessage.textContent = "Bitte warten, Ticket wird generiert...";

  // Formular-Daten aus den (ggf. vom Benutzer geänderten) Feldern erfassen
  const formData = {
    titel: document.getElementById("titel").value.trim(),
    beschreibung: document.getElementById("beschreibung").value.trim(),
    verantwortlicher: document.getElementById("verantwortlicher").value.trim(),
    ansprechpartner: document.getElementById("ansprechpartner").value.trim(),
    ticketTyp: document.getElementById("ticketTyp").value,
    ticketKategorie: document.getElementById("ticketKategorie").value,
    prioritaet: document.getElementById("prioritaet").value,
  };

  // Erneut E-Mail-Daten extrahieren
  const item = Office.context.mailbox.item;
  const subject = item.subject || "";
  const sender = item.from && item.from.emailAddress ? item.from.emailAddress : "Unbekannt";
  const recipients = item.to || [];
  const recipient = recipients.length > 0 ? recipients[0].emailAddress : "Unbekannt";
  let emailBody = "";
  try {
    emailBody = await getEmailBodyAsync();
  } catch (error) {
    console.error("Fehler beim Abrufen des E-Mail-Bodys:", error);
    statusMessage.textContent = "Fehler beim Abrufen des E-Mail-Bodys.";
    return;
  }

  // Prompt erstellen: Hinweis, dass der E-Mail-Verlauf im Feld "beschreibung" enthalten ist
  const prompt = createPrompt(subject, sender, recipient, emailBody, formData);
  console.log("Generierter Prompt:\n", prompt);

  try {
    const aiResponse = await callAiService(prompt);
    console.log("Antwort des AI-Service:", aiResponse);
    const ticketData = JSON.parse(aiResponse);

    // Formularfelder mit den AI-Daten aktualisieren (ggf. Fallback zu den bisherigen Werten)
    document.getElementById("titel").value = ticketData.titel || formData.titel;
    document.getElementById("beschreibung").value = ticketData.beschreibung || formData.beschreibung;
    document.getElementById("verantwortlicher").value = ticketData.verantwortlicher || formData.verantwortlicher;
    document.getElementById("ansprechpartner").value = ticketData.ansprechpartner || formData.ansprechpartner;
    document.getElementById("ticketTyp").value = ticketData.ticketTyp || formData.ticketTyp;
    document.getElementById("ticketKategorie").value = ticketData.ticketKategorie || formData.ticketKategorie;
    document.getElementById("prioritaet").value = ticketData.prioritaet || formData.prioritaet;

    statusMessage.textContent = "Ticket wurde erfolgreich generiert und aktualisiert.";
  } catch (err) {
    console.error("Fehler beim Verarbeiten der AI-Antwort:", err);
    statusMessage.textContent = "Fehler beim Verarbeiten der AI-Antwort.";
  }
}

function createPrompt(subject, sender, recipient, emailBody) {
  return (
    "Erstelle ein Ticket basierend auf den folgenden Informationen.\n" +
    "------------------------------\n" +
    "E-Mail Betreff: " +
    subject +
    "\n" +
    "E-Mail Absender (Ansprechpartner/Kunde): " +
    sender +
    "\n" +
    "E-Mail Empfänger (Verantwortlicher): " +
    recipient +
    "\n" +
    "E-Mail Verlauf (Beschreibung): " +
    emailBody +
    "\n" +
    "------------------------------\n" +
    "Bitte antworte kurz und bündig als JSON mit den Schlüsseln: " +
    "titel, beschreibung, verantwortlicher, ansprechpartner, ticketTyp, ticketKategorie, prioritaet. " +
    "Du bist ein äußerst präziser und hilfreicher Ticket-Assistent. Der Schlüssel 'E-Mail Verlauf (Beschreibung)' enthält den kompletten E-Mail-Verlauf, der möglicherweise HTML-formatiert ist. Entferne alle HTML-Tags und verarbeite den bereinigten Text. Fülle die Felder 'titel', 'beschreibung', 'verantwortlicher', 'ansprechpartner', 'ticketTyp', 'ticketKategorie' und 'prioritaet' vollständig mit den Infos aus dem Mail-Verlauf aus. Falls ein Feld bereits vom Benutzer vorgegeben ist versuch es zu verbessern. Antworte ausschließlich in reinem JSON-Format, ohne Erklärungen oder zusätzliche Texte."
  );
}
