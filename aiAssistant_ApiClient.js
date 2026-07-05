/* =========================
   API / PARSING
========================= */

function fetchAiResponse(messagesArray) {
  const scriptProps = PropertiesService.getScriptProperties();

  const apiKey = scriptProps.getProperty('AI_API_KEY');
  if (!apiKey) throw new Error("AI_API_KEY not found in Script Properties.");

  const url = scriptProps.getProperty('AI_API_ENDPOINT');
  if (!url) throw new Error("AI_API_ENDPOINT not found in Script Properties.");

  const model = scriptProps.getProperty('AI_MODEL');
  if (!model) throw new Error("AI_MODEL not found in Script Properties.");

  const payload = {
    model: model,
    messages: messagesArray,
    temperature: 0.1
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const httpCode = response.getResponseCode();
  const raw = response.getContentText();

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error("API returned non-JSON response. HTTP " + httpCode + ". Raw: " + raw);
  }

  parsed._meta = {
    httpCode: httpCode,
    requestBytes: JSON.stringify(payload).length,
    responseBytes: raw.length
  };

  return parsed;
}

function extractMessageContent(content) {
  if (typeof content === 'string') return content;

  if (Array.isArray(content)) {
    return content.map(part => {
      if (typeof part === 'string') return part;
      if (part && typeof part.text === 'string') return part.text;
      if (part && part.type === 'text' && typeof part.content === 'string') return part.content;
      return '';
    }).join('\n').trim();
  }

  if (content && typeof content.text === 'string') return content.text;

  return String(content || '');
}

function parseAiResponse(json) {
  if (json.error) {
    return {
      message: "API Error: " + (json.error.message || JSON.stringify(json.error)),
      edits: [],
      usage: json.usage || json._meta || {}
    };
  }

  if (!json.choices || !json.choices[0] || !json.choices[0].message) {
    return {
      message: "Unexpected API format.",
      edits: [],
      usage: json.usage || json._meta || {}
    };
  }

  let aiRawReply = extractMessageContent(json.choices[0].message.content).trim();
  let aiResponseObj = { message: "", edits: [] };

  const jsonStartIndex = aiRawReply.indexOf('{');
  const jsonEndIndex = aiRawReply.lastIndexOf('}');

  if (jsonStartIndex !== -1 && jsonEndIndex !== -1 && jsonEndIndex >= jsonStartIndex) {
    try {
      aiResponseObj = JSON.parse(aiRawReply.substring(jsonStartIndex, jsonEndIndex + 1));
    } catch (e) {
      aiResponseObj = {
        message: "JSON formatting error. Raw text: " + aiRawReply,
        edits: []
      };
    }
  } else {
    aiResponseObj = {
      message: aiRawReply,
      edits: []
    };
  }

  if (!Array.isArray(aiResponseObj.edits)) aiResponseObj.edits = [];
  if (typeof aiResponseObj.message !== 'string') aiResponseObj.message = String(aiResponseObj.message || '');
  aiResponseObj.usage = json.usage || json._meta || {};

  return aiResponseObj;
}
