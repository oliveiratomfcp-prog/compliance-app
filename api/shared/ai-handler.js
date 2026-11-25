// Handlers des Functions IA du Tracker CPL :
//   POST /api/ai-search   { query, catalogue }                       -> { resultats: [{ ref, score, pourquoi }] }
//   POST /api/ai-similar  { objet, type, description, catalogue }    -> { similaires: [{ ref, similarite, explication }] }
// Authentification Microsoft obligatoire, limite d'appels par utilisateur, entrées nettoyées,
// sorties validées contre le catalogue reçu. Le contenu des demandes n'est jamais journalisé.
'use strict';

const fs = require('fs');
const path = require('path');
const { authenticate, sendJson, unauthorized } = require('./auth');
const logic = require('./ai-logic');
const { readConfig } = require('./ai-config');
const openai = require('./ai-openai');

const allow = logic.createRateLimiter(15, 60 * 1000);   // 15 appels par minute et par utilisateur

let promptsCache = null;
function readPrompts() {
  if (!promptsCache) promptsCache = logic.splitPrompts(fs.readFileSync(path.join(__dirname, 'prompts-tracker-ai.md'), 'utf8'));
  return promptsCache;
}

function contentLength(req) {
  const headers = (req && req.headers) || {};
  const key = Object.keys(headers).find(function (k) { return k.toLowerCase() === 'content-length'; });
  return key ? parseInt(headers[key], 10) || 0 : 0;
}

const KINDS = {
  search: {
    label: 'recherche',
    step: function (cfg) { return cfg.recherche; },
    validate: logic.validateSearchPayload,
    request: function (prompts, p) { return logic.buildSearchRequest(prompts, p.query, p.catalogue); },
    schemaName: 'resultats_recherche',
    schema: logic.SEARCH_SCHEMA,
    finish: function (json, refs) { return { resultats: logic.validateSearchResult(json, refs) }; }
  },
  similar: {
    label: 'similaires',
    step: function (cfg) { return cfg.similaires; },
    validate: logic.validateSimilarPayload,
    request: function (prompts, p) { return logic.buildSimilarRequest(prompts, p, p.catalogue); },
    schemaName: 'demandes_similaires',
    schema: logic.SIMILAR_SCHEMA,
    finish: function (json, refs) { return { similaires: logic.validateSimilarResult(json, refs) }; }
  }
};

// deps : { fetch, now } injectables pour les tests
function makeHandler(kindName, deps) {
  const kind = KINDS[kindName];
  return async function (context, req) {
    const user = await authenticate(req);
    if (!user) { unauthorized(context); return; }

    if (!allow(user.email)) {
      sendJson(context, 429, { error: { message: 'Trop de recherches IA, réessayez dans une minute.', code: 'rate_limit' } });
      return;
    }

    let cfg;
    let prompts;
    try {
      cfg = readConfig();
      prompts = readPrompts();
    } catch (e) {
      context.log.error('IA ' + kind.label + ' : configuration invalide : ' + e.message);
      sendJson(context, 500, { error: { message: 'Fonction IA non configurée sur le serveur.', code: 'config' } });
      return;
    }

    if (contentLength(req) > logic.LIMITS.maxBodyBytes) {
      sendJson(context, 413, { error: { message: 'Requête trop volumineuse.', code: 'too_large' } });
      return;
    }

    let payload;
    try {
      payload = kind.validate(req.body);
    } catch (e) {
      sendJson(context, e.status || 400, { error: { message: e.message, code: 'bad_request' } });
      return;
    }

    const started = Date.now();
    try {
      const step = kind.step(cfg);
      const out = await openai.respondJson(cfg, step, kind.request(prompts, payload), kind.schemaName, kind.schema, deps);
      const body = kind.finish(out.json, logic.refSet(payload.catalogue));
      const ms = Date.now() - started;
      context.log('IA ' + kind.label + ' : ' + payload.catalogue.length + ' demandes, ' + (body.resultats || body.similaires).length +
        ' résultats, ' + ms + ' ms, tokens ' + out.usage.entree + '/' + out.usage.sortie + ' (cache ' + out.usage.cache + ')');
      body.usage = out.usage;
      body.modele = step.model;
      body.ms = ms;
      sendJson(context, 200, body);
    } catch (e) {
      const status = e && e.status ? e.status : 502;
      context.log.error('IA ' + kind.label + ' : ' + (e && e.code ? e.code : 'erreur') + ' : ' + (e && (e.detail || e.message)));
      sendJson(context, status, { error: { message: e && e.name === 'AiError' ? e.message : 'Service IA indisponible.', code: (e && e.code) || 'ai_error' } });
    }
  };
}

module.exports = { search: makeHandler('search'), similar: makeHandler('similar'), makeHandler: makeHandler };
