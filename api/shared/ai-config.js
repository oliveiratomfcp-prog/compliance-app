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
//   OPENAI_TIMEOUT_MS_TRACKER   délai maximal de la recherche IA en ms (défaut 7000, de 1000 à 40000)
//   OPENAI_TIMEOUT_MS_SIMILAIRES délai maximal des demandes similaires en ms (défaut 15000, de 1000 à 40000)
//   (le navigateur attend un peu plus : 8 s pour la recherche, 16 s pour les demandes similaires)
// Le navigateur ne choisit jamais le modèle : seuls les modèles configurés ici sont utilisés.
'use strict';

const base = require('./docreview-config');

const DEFAULT_EFFORT = 'low';
const DEFAULT_TIMEOUT_MS = 7000;              // recherche IA
const DEFAULT_TIMEOUT_SIMILAIRES_MS = 15000;  // demandes similaires (résumé à rédiger)

function readTimeout(env, name, def) {
  if (!env[name]) return def;
  const v = parseInt(env[name], 10);
  if (!(v >= 1000 && v <= 40000)) throw new Error(name + ' invalide (de 1000 à 40000).');
  return v;
}

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

  const timeoutMs = readTimeout(env, 'OPENAI_TIMEOUT_MS_TRACKER', DEFAULT_TIMEOUT_MS);
  recherche.timeoutMs = timeoutMs;
  similaires.timeoutMs = readTimeout(env, 'OPENAI_TIMEOUT_MS_SIMILAIRES', DEFAULT_TIMEOUT_SIMILAIRES_MS);

  return { apiKey: b.apiKey, baseUrl: b.baseUrl, recherche: recherche, similaires: similaires, timeoutMs: timeoutMs };
}

module.exports = { readConfig: readConfig, DEFAULT_EFFORT: DEFAULT_EFFORT, DEFAULT_TIMEOUT_MS: DEFAULT_TIMEOUT_MS, DEFAULT_TIMEOUT_SIMILAIRES_MS: DEFAULT_TIMEOUT_SIMILAIRES_MS };
