
// Para poder tener IA

function chat() {
    return HtmlService.createHtmlOutputFromFile('chat').getContent();
}

function doPost(data) {
    try {
        const userMessage = data.message || '';

        if (!userMessage) {
            return { reply: "Mensaje vacío." };
        }

        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME3);
        const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();

        const prompt = `Usa la siguiente tabla para responder:\n${JSON.stringify(values)}\nPregunta: ${userMessage}`;

        const apiKey = PropertiesService.getScriptProperties().getProperty('API_KEY');// Hay que guardar el codigo en proipedades de secuencias de comandos 
        Logger.log("API Key obtenida: " + apiKey);
        const url = `https://generativelanguage.googleapis.com/v1/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

        const payload = { contents: [{ parts: [{ text: prompt }] }] };
        const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload),muteHttpExceptions: true };
        

        const response = UrlFetchApp.fetch(url, options);
        Logger.log("Respuesta completa: " + response.getContentText());
        const result = JSON.parse(response.getContentText());

        const reply = result?.candidates?.[0]?.content?.parts?.[0]?.text || "No se pudo generar respuesta.";

        return { reply: reply };

    } catch (error) {
        Logger.log("Error en doPost: " + error);
        return { reply: "Error interno en el servidor." };
    }
}
