// Appels HTTP à l'API Responses d'OpenAI (création en mode background, lecture, annulation).
'use strict';

const fetch = require('node-fetch');

async function call(cfg, method, path, body) {
  let res;
  try {
    res = await fetch(cfg.baseUrl + path, {
      method: method,
      headers: { 'Authorization': 'Bearer ' + cfg.apiKey, 'Content-Type': 'application/json' },
      body: body ? JSON.stringify(body) : undefined,
      timeout: 30000
    });
  } catch (e) {
    throw new Error('OpenAI injoignable : ' + e.message);
  }
  let data = null;
  try { data = await res.json(); } catch (e) { data = null; }
  if (!res.ok) {
    const msg = data && data.error && data.error.message ? data.error.message : 'HTTP ' + res.status;
    const err = new Error('OpenAI : ' + msg);
    err.status = res.status;
    throw err;
  }
  return data;
}

function createResponse(cfg, body) { return call(cfg, 'POST', '/responses', body); }
function getResponse(cfg, id) { return call(cfg, 'GET', '/responses/' + encodeURIComponent(id)); }
function cancelResponse(cfg, id) { return call(cfg, 'POST', '/responses/' + encodeURIComponent(id) + '/cancel'); }

module.exports = { createResponse: createResponse, getResponse: getResponse, cancelResponse: cancelResponse };
