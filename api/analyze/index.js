const fetch = require('node-fetch');
const { authenticate, sendJson, unauthorized } = require('../shared/auth');

// Relais OpenAI sécurisé (transition) : utilisateur Eiffel authentifié obligatoire,
// modèles limités à une liste blanche, corps réduit aux seuls champs attendus.
const ALLOWED_MODELS = ['gpt-4o', 'gpt-4o-mini'];
const MAX_MESSAGES = 4;
const MAX_CONTENT_CHARS = 300000;

function sanitizeBody(body) {
  if (!body || typeof body !== 'object') throw new Error('Requête invalide.');
  if (ALLOWED_MODELS.indexOf(body.model) === -1) throw new Error('Modèle non autorisé.');
  if (!Array.isArray(body.messages) || body.messages.length === 0 || body.messages.length > MAX_MESSAGES) {
    throw new Error('Messages invalides.');
  }
  const messages = body.messages.map(function (m) {
    if (!m || (m.role !== 'user' && m.role !== 'system') || typeof m.content !== 'string' || m.content.length > MAX_CONTENT_CHARS) {
      throw new Error('Message invalide.');
    }
    return { role: m.role, content: m.content };
  });
  const clean = { model: body.model, messages: messages };
  if (typeof body.temperature === 'number' && body.temperature >= 0 && body.temperature <= 2) clean.temperature = body.temperature;
  if (body.response_format && body.response_format.type === 'json_object') clean.response_format = { type: 'json_object' };
  return clean;
}

module.exports = async function (context, req) {
  const user = await authenticate(req);
  if (!user) { unauthorized(context); return; }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    sendJson(context, 500, { error: { message: 'OPENAI_API_KEY non configurée sur le serveur.' } });
    return;
  }

  let clean;
  try {
    clean = sanitizeBody(req.body);
  } catch (e) {
    sendJson(context, 400, { error: { message: e.message } });
    return;
  }

  const baseUrl = String(process.env.OPENAI_BASE_URL || 'https://api.openai.com/v1').replace(/\/+$/, '');
  try {
    const response = await fetch(baseUrl + '/chat/completions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + apiKey },
      body: JSON.stringify(clean)
    });
    const data = await response.json();
    sendJson(context, response.status, data);
  } catch (e) {
    sendJson(context, 502, { error: { message: e.message } });
  }
};
