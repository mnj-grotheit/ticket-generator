/* eslint-disable no-undef */
import { ai_config } from "../../ai_config.js";
import { GoogleGenerativeAI } from "@google/generative-ai";

const apiKey = ai_config.GEMINI_API_KEY;
const genAI = new GoogleGenerativeAI(apiKey);

export async function callAiService(prompt) {
  try {
    // Konfiguriere das Modell mit erweiterter Systeminstruktion
    const model = genAI.getGenerativeModel({
      model: "gemini-2.0-flash",
      systemInstruction:
        "Du bist ein äußerst präziser und hilfreicher Ticket-Assistent. Der Schlüssel 'E-Mail Verlauf (Beschreibung)' enthält den kompletten E-Mail-Verlauf, der möglicherweise HTML-formatiert ist. Entferne alle HTML-Tags und verarbeite den bereinigten Text. Fülle die Felder 'titel', 'beschreibung', 'verantwortlicher', 'ansprechpartner', 'ticketTyp', 'ticketKategorie' und 'prioritaet' vollständig mit den Infos aus dem Mail-Verlauf aus. Falls ein Feld bereits vom Benutzer vorgegeben ist versuch es zu verbessern. Antworte ausschließlich in reinem JSON-Format, ohne Erklärungen oder zusätzliche Texte.",
    });

    // Konfiguration der Generierung, inklusive Response-Schema
    const generationConfig = {
      temperature: 1,
      topP: 0.95,
      topK: 40,
      maxOutputTokens: 8192,
      responseMimeType: "application/json", // JSON als Antwort
      responseSchema: {
        type: "object",
        properties: {
          titel: { type: "string" },
          beschreibung: { type: "string" },
          verantwortlicher: { type: "string" },
          ansprechpartner: { type: "string" },
          ticketTyp: { type: "string" },
          ticketKategorie: { type: "string" },
          prioritaet: { type: "string" },
        },
        required: [
          "titel",
          "beschreibung",
          "verantwortlicher",
          "ansprechpartner",
          "ticketTyp",
          "ticketKategorie",
          "prioritaet",
        ],
      },
    };

    // Starte eine Chat-Session
    const chatSession = model.startChat({
      generationConfig,
      history: [],
    });

    // Sende den Prompt und warte auf die Antwort
    const result = await chatSession.sendMessage(prompt);
    const responseText = result.response.text();
    console.log("Gemini AI response:", responseText);

    return responseText;
  } catch (error) {
    console.error("Fehler bei der Anfrage an den Gemini AI-Service:", error);
    throw error;
  }
}
