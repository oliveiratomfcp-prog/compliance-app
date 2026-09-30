// POST /api/docreview-start : démarre une étape d'analyse (lot de pages ou synthèse)
// en mode background de l'API Responses, et renvoie l'identifiant de la réponse
// avec un ticket lié à l'utilisateur. Le prompt est construit ici, côté serveur.
'use strict';

const { authenticate, sendJson, unauthorized } = require('../shared/auth');
const logic = require('../shared/docreview-logic');
const { readConfig, readRules } = require('../shared/docreview-config');
const openai = require('../shared/openai');
const { signTicket } = require('../shared/ticket');

module.exports = async function (context, req) {
  const user = await authenticate(req);
  if (!user) { unauthorized(context); return; }

  let cfg;
  try {
    cfg = readConfig();
  } catch (e) {
    sendJson(context, 500, { error: { message: 'Configuration serveur invalide : ' + e.message } });
    return;
  }
  if (!cfg.apiKey) {
    sendJson(context, 500, { error: { message: 'OPENAI_API_KEY non configurée sur le serveur.' } });
    return;
  }

  let payload;
  try {
    payload = logic.validateStartPayload(req.body, logic.LIMITS);
  } catch (e) {
    sendJson(context, 400, { error: { message: e.message } });
    return;
  }

  let request;
  try {
    request = logic.buildOpenAIRequest(payload, readRules(), cfg);
  } catch (e) {
    context.log.error('Doc Review : construction du prompt impossible', e.message);
    sendJson(context, 500, { error: { message: 'Règles d\'analyse illisibles : ' + e.message } });
    return;
  }

  try {
    const resp = await openai.createResponse(cfg, request.body);
    if (!resp || !logic.isValidResponseId(resp.id)) throw new Error('Réponse OpenAI sans identifiant.');
    sendJson(context, 200, { id: resp.id, ticket: signTicket(resp.id, user.email), status: resp.status || 'queued', modele: request.model });
  } catch (e) {
    context.log.error('Doc Review : démarrage impossible', e.message);
    sendJson(context, 502, { error: { message: e.message } });
  }
};
