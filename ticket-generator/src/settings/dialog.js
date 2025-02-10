/* eslint-disable no-undef */
/* global Office, $, fetch */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // eslint-disable-next-line no-undef
    console.log("Outlook Add-In ist bereit.");

    // Event-Listener für den Formular-Submit
    $("#ticketForm").on("submit", function (e) {
      e.preventDefault(); // Standard-Submit verhindern

      // 1. Formular-Daten auslesen
      const formData = {
        titel: $("#titel").val(),
        beschreibung: $("#beschreibung").val(),
        verantwortlicher: $("#verantwortlicher").val(),
        ansprechpartner: $("#ansprechpartner").val(),
        ticketTyp: $("#ticketTyp").val(),
        ticketKategorie: $("#ticketKategorie").val(),
        prioritaet: $("#prioritaet").val(),
      };

      // 2. Outlook Mail-Daten extrahieren
      const item = Office.context.mailbox.item;
      const subject = item.subject;
      const sender = item.from && item.from.emailAddress ? item.from.emailAddress : "Unbekannt";

      // Den Body der E-Mail asynchron abrufen
      item.body.getAsync("text", function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const emailBody = asyncResult.value;

          // 3. Einen sinnvollen Prompt erstellen, der alle Informationen kombiniert
          let prompt = "Erstelle ein Ticket basierend auf den folgenden Informationen:\n";
          prompt += "------------------------------\n";
          prompt += "E-Mail Betreff: " + subject + "\n";
          prompt += "E-Mail Absender: " + sender + "\n";
          prompt += "E-Mail Body: " + emailBody + "\n";
          prompt += "------------------------------\n";
          prompt += "Zusätzliche Ticket-Informationen:\n";
          prompt += "Titel: " + formData.titel + "\n";
          prompt += "Beschreibung: " + formData.beschreibung + "\n";
          prompt += "Verantwortlicher: " + formData.verantwortlicher + "\n";
          prompt += "Ansprechpartner (Kunde): " + formData.ansprechpartner + "\n";
          prompt += "Ticket-Typ: " + formData.ticketTyp + "\n";
          prompt += "Ticket-Kategorie: " + formData.ticketKategorie + "\n";
          prompt += "Priorität: " + formData.prioritaet + "\n";
          prompt += "------------------------------\n";
          prompt += "Bitte formatiere die Antwort als JSON mit folgenden Schlüsseln: ";
          prompt += "titel, beschreibung, verantwortlicher, ansprechpartner, ticketTyp, ticketKategorie, prioritaet.";

          console.log("An AI Service gesendeter Prompt:\n", prompt);

          // 4. Authentifizieren und den Prompt an den AI Service senden (Platzhalter!)
          fetch("https://api.openai.com/v1/chat/completions", {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: "Bearer YOUR_API_KEY", // Ersetze diesen Platzhalter mit Deinem echten API-Key
            },
            body: JSON.stringify({
              model: "gpt-3.5-turbo",
              messages: [
                { role: "system", content: "Du bist ein hilfreicher Ticket-Assistent." },
                { role: "user", content: prompt },
              ],
              temperature: 0.7,
              max_tokens: 500,
            }),
          })
            .then((response) => response.json())
            .then((data) => {
              console.log("Antwort des AI Service:", data);

              // 5. Die Antwort der KI verarbeiten
              let aiResponse = "";
              if (data && data.choices && data.choices.length > 0) {
                aiResponse = data.choices[0].message.content;
              } else {
                console.error("Keine gültige Antwort erhalten.");
                return;
              }

              // Die AI-Antwort sollte ein JSON-Objekt enthalten. Versuche, dieses zu parsen.
              try {
                const ticketData = JSON.parse(aiResponse);

                // 6. Formularfelder mit den Daten aus der KI-Antwort befüllen
                $("#titel").val(ticketData.titel || formData.titel);
                $("#beschreibung").val(ticketData.beschreibung || formData.beschreibung);
                $("#verantwortlicher").val(ticketData.verantwortlicher || formData.verantwortlicher);
                $("#ansprechpartner").val(ticketData.ansprechpartner || formData.ansprechpartner);
                $("#ticketTyp").val(ticketData.ticketTyp || formData.ticketTyp);
                $("#ticketKategorie").val(ticketData.ticketKategorie || formData.ticketKategorie);
                $("#prioritaet").val(ticketData.prioritaet || formData.prioritaet);

                console.log("Formular wurde mit den AI-Daten aktualisiert.");
              } catch (err) {
                console.error("Fehler beim Parsen der AI-Antwort:", err);
              }
            })
            .catch((error) => {
              console.error("Fehler bei der Kommunikation mit dem AI-Service:", error);
            });
        } else {
          console.error("Fehler beim Abrufen des E-Mail-Bodys:", asyncResult.error);
        }
      });
    });
  }
});
