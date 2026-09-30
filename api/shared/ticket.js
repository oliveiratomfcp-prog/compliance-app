// Ticket liant une réponse OpenAI à l'utilisateur qui l'a lancée : sans ce ticket,
// un autre utilisateur ne peut ni lire ni annuler l'analyse, même s'il connaît l'identifiant.
// HMAC-SHA256 avec une clé dérivée de OPENAI_API_KEY (aucun secret supplémentaire à gérer).
'use strict';

const crypto = require('crypto');

function key() {
  return crypto.createHash('sha256').update('docreview-ticket-v1:' + String(process.env.OPENAI_API_KEY || '')).digest();
}

function signTicket(id, email) {
  return crypto.createHmac('sha256', key()).update(String(id) + '|' + String(email)).digest('hex');
}

function verifyTicket(id, email, ticket) {
  if (typeof ticket !== 'string' || !/^[0-9a-f]{64}$/.test(ticket)) return false;
  const expected = Buffer.from(signTicket(id, email), 'hex');
  const given = Buffer.from(ticket, 'hex');
  return expected.length === given.length && crypto.timingSafeEqual(expected, given);
}

module.exports = { signTicket: signTicket, verifyTicket: verifyTicket };
