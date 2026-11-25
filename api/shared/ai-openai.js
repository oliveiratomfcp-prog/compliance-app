// Appel synchrone à l'API Responses d'OpenAI avec sortie JSON stricte (json_schema), comme
// Doc Review mais sans le mode background : la réponse doit tenir dans le délai configuré.
// Les erreurs sont normalisées en AiError : le statut HTTP renvoyé au navigateur et un code
// court, sans détail technique (les détails vont uniquement dans les journaux serveur).
'use strict';

const { extractOutput } = require('./docreview-logic');

// Raisonnement et réponse compris (une réponse tronquée est une erreur, jamais un résultat partiel)
const MAX_OUTPUT_TOKENS = 12000;

class AiError extends Error {
  constructor(message, status, code, detail) {
    super(message);
    this.name = 'AiError';
    this.status = status || 502;
    this.code = code || 'ai_error';
    this.detail = detail || '';
  }
}

// request : { instructions, input } ; step : { model, effort } ; deps.fetch injectable pour les tests.
async function respondJson(cfg, step, request, schemaName, schema, deps) {
  const fetchImpl = (deps && deps.fetch) || require('node-fetch');
  const body = {
    model: step.model,
    instructions: request.instructions,
    input: request.input,
    reasoning: { effort: step.effort },
    text: { format: { type: 'json_schema', name: schemaName, schema: schema, strict: true } },
    store: false,
    max_output_tokens: MAX_OUTPUT_TOKENS
  };

  const controller = new AbortController();
  const timer = setTimeout(function () { controller.abort(); }, step.timeoutMs || cfg.timeoutMs);
  let res;
  let text;
  try {
    res = await fetchImpl(cfg.baseUrl + '/responses', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + cfg.apiKey, 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: controller.signal
    });
    text = await res.text();
  } catch (e) {
    if (e && e.name === 'AbortError') throw new AiError('Le modèle a mis trop de temps à répondre.', 504, 'timeout');
    throw new AiError('Service IA injoignable.', 502, 'unreachable', e && e.message);
  } finally {
    clearTimeout(timer);
  }

  let data = null;
  try { data = JSON.parse(text); } catch (e) { data = null; }

  if (!res.ok) {
    const apiErr = data && data.error ? data.error : {};
    const code = String(apiErr.code || 'http_' + res.status);
    const status = res.status === 429 ? 429 : 502;
    throw new AiError(res.status === 429 ? 'Service IA saturé, réessayez dans un instant.' : 'Service IA indisponible.', status, code, apiErr.message);
  }
  if (!data) throw new AiError('Réponse IA illisible.', 502, 'invalid_json');
  if (data.status === 'failed') throw new AiError('Service IA indisponible.', 502, 'failed', data.error && data.error.message);
  if (data.status === 'incomplete') throw new AiError('Réponse IA tronquée.', 502, 'truncated', data.incomplete_details && data.incomplete_details.reason);

  const out = extractOutput(data);
  if (out.refusal) throw new AiError('Requête refusée par le modèle.', 502, 'refusal', out.refusal);
  if (!out.text) throw new AiError('Réponse IA vide.', 502, 'empty');

  let json;
  try { json = JSON.parse(out.text); } catch (e) { throw new AiError('Réponse IA illisible.', 502, 'invalid_json'); }

  const u = data.usage || {};
  return {
    json: json,
    usage: {
      entree: u.input_tokens || 0,
      sortie: u.output_tokens || 0,
      cache: (u.input_tokens_details && u.input_tokens_details.cached_tokens) || 0
    }
  };
}

module.exports = { respondJson: respondJson, AiError: AiError, MAX_OUTPUT_TOKENS: MAX_OUTPUT_TOKENS };
