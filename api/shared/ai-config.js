// Configuration des Functions IA du Tracker CPL (recherche IA, demandes similaires).
// Elles réutilisent la configuration OpenAI de Doc Review : même clé, même URL de base, et par
// défaut le même modèle. Aucun paramètre supplémentaire n'est donc obligatoire.
//   OPENAI_API_KEY              clé de l'API OpenAI (obligatoire, déjà configurée pour Doc Review)
//   OPENAI_BASE_URL             URL de base (défaut https://api.openai.com/v1, voir Doc Review)
// Réglages facultatifs, propres au Tracker :
//   OPENAI_MODEL_RECHERCHE      modèle de la recherche IA (défaut : OPENAI_MODEL_SYNTHESE, sinon le modèle par défaut de Doc Review)
//   OPENAI_MODEL_SIMILAIRES     modèle des demandes similaires (défaut : celui de la recherche)
//   OPENAI_EFFORT_RECHERCHE     effort de raisonnement de la recherche (défaut low)
//   OPENAI_EFFORT_SIMILAIRES    effort de raisonnement des demandes similaires (défaut : celui de la recherche)
//   OPENAI_TIMEOUT_MS_TRACKER   délai maximal d'un appel au modèle en ms (défaut 7000, de 1000 à 40000)
// Le navigateur ne choisit jamais le modèle : seuls les modèles configurés ici sont utilisés.
'use strict';

const base = require('./docreview-config');

const DEFAULT_EFFORT = 'low';
const DEFAULT_TIMEOUT_MS = 7000;

function readConfig(env) {
  env = env || process.env;
  const b = base.readBase(env);
  if (!b.apiKey) throw new Error('OPENAI_API_KEY non configurée.');

  const defaultModel = env.OPENAI_MODEL_SYNTHESE || base.DEFAULT_MODEL;
  const recherche = {
    model: base.model(env.OPENAI_MODEL_RECHERCHE || defaultModel, 'OPENAI_MODEL_RECHERCHE'),
    effort: base.effort(env.OPENAI_EFFORT_RECHERCHE, DEFAULT_EFFORT, 'OPENAI_EFFORT_RECHERCHE')
  };
  const similaires = {
    model: base.model(env.OPENAI_MODEL_SIMILAIRES || recherche.model, 'OPENAI_MODEL_SIMILAIRES'),
    effort: base.effort(env.OPENAI_EFFORT_SIMILAIRES, recherche.effort, 'OPENAI_EFFORT_SIMILAIRES')
  };

  let timeoutMs = DEFAULT_TIMEOUT_MS;
  if (env.OPENAI_TIMEOUT_MS_TRACKER) {
    timeoutMs = parseInt(env.OPENAI_TIMEOUT_MS_TRACKER, 10);
    if (!(timeoutMs >= 1000 && timeoutMs <= 40000)) throw new Error('OPENAI_TIMEOUT_MS_TRACKER invalide (de 1000 à 40000).');
  }

  return { apiKey: b.apiKey, baseUrl: b.baseUrl, recherche: recherche, similaires: similaires, timeoutMs: timeoutMs };
}

module.exports = { readConfig: readConfig, DEFAULT_EFFORT: DEFAULT_EFFORT, DEFAULT_TIMEOUT_MS: DEFAULT_TIMEOUT_MS };
