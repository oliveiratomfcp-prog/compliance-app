// Configuration des Functions Doc Review, lue dans les Application Settings Azure.
//   OPENAI_API_KEY          clé de l'API OpenAI (obligatoire)
//   OPENAI_BASE_URL         URL de base (défaut https://api.openai.com/v1 ; UE : https://eu.api.openai.com/v1)
//   OPENAI_MODEL_PAGES      modèle de l'étape 1, analyse par lots de pages (défaut gpt-5.6-terra)
//   OPENAI_MODEL_SYNTHESE   modèle de l'étape 2, synthèse (défaut gpt-5.6-terra)
//   OPENAI_EFFORT_PAGES     effort de raisonnement de l'étape 1 (défaut low)
//   OPENAI_EFFORT_SYNTHESE  effort de raisonnement de l'étape 2 (défaut medium)
// Le navigateur ne choisit jamais le modèle : seuls les modèles configurés ici sont utilisés.
'use strict';

const fs = require('fs');
const path = require('path');
const logic = require('./docreview-logic');

const DEFAULT_MODEL = 'gpt-5.6-terra';
const EFFORTS = ['none', 'minimal', 'low', 'medium', 'high', 'xhigh', 'max'];

// Budget de sortie (raisonnement + réponse) par appel. OpenAI recommande de réserver
// au moins 25 000 tokens pour les modèles de raisonnement.
const MAX_OUTPUT_TOKENS_PAGES = 32000;
const MAX_OUTPUT_TOKENS_SYNTHESE = 64000;

function model(value, name) {
  const v = String(value || DEFAULT_MODEL).trim();
  if (!/^[A-Za-z0-9._:-]{1,100}$/.test(v)) throw new Error(name + ' invalide.');
  return v;
}

function effort(value, fallback, name) {
  const v = String(value || fallback).trim().toLowerCase();
  if (EFFORTS.indexOf(v) === -1) throw new Error(name + ' invalide (valeurs : ' + EFFORTS.join(', ') + ').');
  return v;
}

function readConfig() {
  const env = process.env;
  const baseUrl = String(env.OPENAI_BASE_URL || 'https://api.openai.com/v1').trim().replace(/\/+$/, '');
  if (!/^https:\/\/[A-Za-z0-9.-]+(\/[A-Za-z0-9._\/-]*)?$/.test(baseUrl)) throw new Error('OPENAI_BASE_URL invalide (https obligatoire).');
  return {
    apiKey: env.OPENAI_API_KEY || '',
    baseUrl: baseUrl,
    pages: {
      model: model(env.OPENAI_MODEL_PAGES, 'OPENAI_MODEL_PAGES'),
      effort: effort(env.OPENAI_EFFORT_PAGES, 'low', 'OPENAI_EFFORT_PAGES'),
      maxOutputTokens: MAX_OUTPUT_TOKENS_PAGES
    },
    synthese: {
      model: model(env.OPENAI_MODEL_SYNTHESE, 'OPENAI_MODEL_SYNTHESE'),
      effort: effort(env.OPENAI_EFFORT_SYNTHESE, 'medium', 'OPENAI_EFFORT_SYNTHESE'),
      maxOutputTokens: MAX_OUTPUT_TOKENS_SYNTHESE
    }
  };
}

let rulesCache = null;
function readRules() {
  if (!rulesCache) rulesCache = logic.splitRules(fs.readFileSync(path.join(__dirname, 'prompts-docreview.md'), 'utf8'));
  return rulesCache;
}

module.exports = { readConfig: readConfig, readRules: readRules, EFFORTS: EFFORTS };
