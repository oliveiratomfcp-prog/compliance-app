// GET /api/docreview-status?id=...&ticket=...&etape=pages|synthese : état d'une étape d'analyse.
// Renvoie { status: pending | completed | failed | cancelled, resultat?, message?, usage? }.
'use strict';

const { authenticate, sendJson, unauthorized } = require('../shared/auth');
const logic = require('../shared/docreview-logic');
const { readConfig } = require('../shared/docreview-config');
const openai = require('../shared/openai');
const { verifyTicket } = require('../shared/ticket');

module.exports = async function (context, req) {
  const user = await authenticate(req);
  if (!user) { unauthorized(context); return; }

  const q = req.query || {};
  const id = q.id;
  const etape = q.etape;
  if (!logic.isValidResponseId(id) || logic.ETAPES.indexOf(etape) === -1) {
    sendJson(context, 400, { error: { message: 'Paramètres invalides.' } });
    return;
  }
  if (!verifyTicket(id, user.email, q.ticket)) {
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
    const resp = await openai.getResponse(cfg, id);
    sendJson(context, 200, logic.interpretResponse(resp, etape));
  } catch (e) {
    context.log.error('Doc Review : lecture de l\'état impossible', e.message);
    sendJson(context, 502, { error: { message: e.message } });
  }
};
