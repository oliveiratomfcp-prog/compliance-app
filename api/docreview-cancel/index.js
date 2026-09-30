// POST /api/docreview-cancel { id, ticket } : annule une étape d'analyse en cours.
'use strict';

const { authenticate, sendJson, unauthorized } = require('../shared/auth');
const logic = require('../shared/docreview-logic');
const { readConfig } = require('../shared/docreview-config');
const openai = require('../shared/openai');
const { verifyTicket } = require('../shared/ticket');

module.exports = async function (context, req) {
  const user = await authenticate(req);
  if (!user) { unauthorized(context); return; }

  const b = req.body || {};
  if (!logic.isValidResponseId(b.id)) {
    sendJson(context, 400, { error: { message: 'Paramètres invalides.' } });
    return;
  }
  if (!verifyTicket(b.id, user.email, b.ticket)) {
    sendJson(context, 403, { error: { message: 'Accès refusé à cette analyse.' } });
    return;
  }

  let cfg;
  try {
    cfg = readConfig();
  } catch (e) {
    sendJson(context, 500, { error: { message: 'Configuration serveur invalide : ' + e.message } });
    return;
  }

  try {
    const resp = await openai.cancelResponse(cfg, b.id);
    sendJson(context, 200, { status: resp && resp.status ? resp.status : 'cancelled' });
  } catch (e) {
    sendJson(context, 502, { error: { message: e.message } });
  }
};
