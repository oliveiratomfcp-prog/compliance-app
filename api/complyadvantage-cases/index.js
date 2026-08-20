// ============================================================================
// Azure Function : GET /api/complyadvantage-cases
// ----------------------------------------------------------------------------
// Déclenchée uniquement par un clic utilisateur côté front (complyadvantage.html,
// bouton "Lancer la collecte"). Pas de timer, pas de webhook.
//
// Fait le pont vers l'API ComplyAdvantage Mesh (cases/alertes AML) en gardant
// les secrets côté serveur, sur le même principe que api/analyze.
//
// Variables d'environnement attendues (Application Settings de la Static Web
// App / Function App côté Azure Portal — JAMAIS dans local.settings.json commité,
// JAMAIS dans le front-end) :
//   - COMPLYADVANTAGE_ACCESS_KEY     : access key Mesh (client_id du flux OAuth2)
//   - COMPLYADVANTAGE_ACCESS_SECRET  : access secret Mesh (client_secret)
//   - COMPLYADVANTAGE_BASE_URL       : (optionnel) URL de base de l'API Mesh.
//       Défaut ci-dessous — voir le point "À CONFIRMER" juste en dessous. Permet
//       de corriger l'URL depuis le portail Azure sans redéployer le code si la
//       valeur par défaut s'avère incorrecte.
//
// ⚠️ À CONFIRMER AVANT MISE EN PROD (cf. cahier des charges, point D) :
//   1. L'URL de base exacte de l'API Mesh et le chemin de l'endpoint de token
//      (cf. https://docs.mesh.complyadvantage.com/reference/createtokenv3).
//      La valeur MESH_BASE_URL_DEFAULT ci-dessous est une hypothèse à vérifier
//      depuis le dashboard Mesh de l'organisation avant tout premier appel réel.
//   2. Le format exact du corps de la requête de token (nom des champs attendus
//      par /v3/token : client_id/client_secret, ou access_key/access_secret,
//      ou autre — à ajuster dans getMeshToken() une fois confirmé).
//   3. Le nom exact du champ "assigné à" dans la réponse réelle de GET /v2/cases
//      (non garanti par la documentation publique). Le mapping ci-dessous teste
//      plusieurs noms plausibles et retombe sur "Non assigné" si aucun ne
//      correspond — à ajuster dans extractAssignee() après un premier appel
//      de test avec les vraies clés API.
// ============================================================================

const fetch = require('node-fetch');

const MESH_BASE_URL_DEFAULT = 'https://api.mesh.complyadvantage.com'; // À CONFIRMER — voir en-tête du fichier
const TOKEN_PATH = '/v3/token'; // cf. docs.mesh.complyadvantage.com/reference/createtokenv3 — À CONFIRMER
const CASES_PATH = '/v2/cases';

const PAGE_SIZE = 100; // taille de page pour la pagination de GET /v2/cases
const MAX_PAGES = 200; // garde-fou anti-boucle infinie (20 000 cas max)

// Si le nombre de cas dépasse ce seuil, on n'appelle PAS systématiquement
// GET /v2/cases/{id}/alerts pour chacun (trop coûteux en volumétrie) : le
// nombre d'alertes est alors renvoyé à null ("à charger") plutôt que compté.
const MAX_ALERTS_AUTO_FETCH = 150;
// Nombre d'appels /alerts menés en parallèle quand on les charge.
const ALERTS_CONCURRENCY = 5;

// Le token Mesh est valide 24h et n'a pas de mécanisme de refresh documenté :
// on le regénère entièrement à l'expiration. Mis en cache en mémoire au niveau
// du module Node pour être réutilisé entre deux invocations tant que l'instance
// de la Function reste chaude (aucune persistance externe, conformément au
// cahier des charges — un cold start régénère simplement un nouveau token).
let cachedToken = null;   // string
let tokenExpiresAt = 0;   // timestamp ms

// Libellés techniques -> lisibles pour case_type (le front-end fait aussi sa
// propre traduction ; ce mapping sert de filet si jamais l'API renvoie une
// valeur non prévue côté front).
const CASE_TYPE_LABELS = {
  CUSTOMER_ONBOARDING: 'Onboarding client',
  CUSTOMER_MONITORING: 'Surveillance client',
  PAYMENT_SCREENING: 'Filtrage paiement',
  TRANSACTION_MONITORING: 'Surveillance transaction'
};

module.exports = async function (context, req) {
  const accessKey = process.env.COMPLYADVANTAGE_ACCESS_KEY;
  const accessSecret = process.env.COMPLYADVANTAGE_ACCESS_SECRET;

  if (!accessKey || !accessSecret) {
    context.res = jsonResponse(500, {
      error: { message: "Configuration serveur manquante : COMPLYADVANTAGE_ACCESS_KEY et/ou COMPLYADVANTAGE_ACCESS_SECRET ne sont pas définies dans les Application Settings de l'Azure Function." }
    });
    return;
  }

  const baseUrl = (process.env.COMPLYADVANTAGE_BASE_URL || MESH_BASE_URL_DEFAULT).replace(/\/+$/, '');

  let token;
  try {
    token = await getMeshToken(baseUrl, accessKey, accessSecret);
  } catch (e) {
    context.log.error('ComplyAdvantage Mesh — échec authentification :', e);
    context.res = jsonResponse(502, {
      error: { message: `Authentification auprès de ComplyAdvantage Mesh impossible : ${e.message}` }
    });
    return;
  }

  let rawCases;
  try {
    rawCases = await fetchAllCases(baseUrl, token);
  } catch (e) {
    context.log.error('ComplyAdvantage Mesh — échec récupération des cas :', e);
    context.res = jsonResponse(502, {
      error: { message: `Récupération des cas ComplyAdvantage impossible : ${e.message}` }
    });
    return;
  }

  const skipAlerts = rawCases.length > MAX_ALERTS_AUTO_FETCH;
  let alertsCounts = new Map();
  if (!skipAlerts) {
    try {
      alertsCounts = await fetchAlertsCounts(baseUrl, token, rawCases.map(c => c.case_identifier));
    } catch (e) {
      // Une défaillance sur le comptage des alertes ne doit pas faire échouer
      // toute la collecte : on continue avec un compteur "à charger" (null).
      context.log.warn('ComplyAdvantage Mesh — comptage des alertes partiellement indisponible :', e.message);
    }
  }

  const now = Date.now();
  const normalized = rawCases.map(c => normalizeCase(c, alertsCounts, skipAlerts, now));

  context.res = jsonResponse(200, {
    cases: normalized,
    meta: {
      total: normalized.length,
      fetchedAt: new Date(now).toISOString(),
      alertsSkipped: skipAlerts // true = volumétrie trop importante pour compter les alertes de chaque cas automatiquement
    }
  });
};

// ----------------------------------------------------------------------------
// Authentification OAuth2 (client credentials)
// ----------------------------------------------------------------------------
async function getMeshToken(baseUrl, accessKey, accessSecret) {
  if (cachedToken && Date.now() < tokenExpiresAt) {
    return cachedToken;
  }

  // Corps de requête confirmé par un appel de test à l'API Mesh : le champ
  // secret s'appelle "secret", pas "access_secret" (une valeur "access_secret"
  // provoque un 400 "access_key is missing or cannot be decoded").
  const response = await fetch(`${baseUrl}${TOKEN_PATH}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      access_key: accessKey,
      secret: accessSecret
    })
  });

  if (!response.ok) {
    const errText = await response.text().catch(() => '');
    throw new Error(`Endpoint de token Mesh (${TOKEN_PATH}) → HTTP ${response.status}${errText ? ` : ${errText.slice(0, 300)}` : ''}`);
  }

  const data = await response.json();
  const accessToken = data.access_token || data.token;
  if (!accessToken) {
    throw new Error("La réponse de l'endpoint de token Mesh ne contient pas de champ access_token/token reconnu.");
  }

  // Durée de vie : 24h annoncées par Mesh. On reprend expires_in si présent
  // (en secondes), sinon on retombe sur 24h moins une marge de sécurité de 5 min.
  const ttlSeconds = Number.isFinite(data.expires_in) ? data.expires_in : (24 * 60 * 60);
  cachedToken = accessToken;
  tokenExpiresAt = Date.now() + Math.max(ttlSeconds - 300, 60) * 1000;

  return cachedToken;
}

// ----------------------------------------------------------------------------
// GET /v2/cases avec pagination — pas de filtre de statut côté API, on récupère
// tout et le filtrage se fait côté front.
// ----------------------------------------------------------------------------
async function fetchAllCases(baseUrl, token) {
  const all = [];
  let pageNumber = 1;

  while (pageNumber <= MAX_PAGES) {
    const url = `${baseUrl}${CASES_PATH}?page_number=${pageNumber}&page_size=${PAGE_SIZE}`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) {
      const errText = await response.text().catch(() => '');
      throw new Error(`GET ${CASES_PATH} (page ${pageNumber}) → HTTP ${response.status}${errText ? ` : ${errText.slice(0, 300)}` : ''}`);
    }

    const data = await response.json();
    // Forme de réponse supposée : { data: [...] } ou directement un tableau —
    // on gère les deux au cas où (à ajuster si la forme réelle diffère).
    const pageItems = Array.isArray(data) ? data : (data.data || data.cases || []);
    all.push(...pageItems);

    if (pageItems.length < PAGE_SIZE) break; // dernière page atteinte
    pageNumber++;
  }

  return all;
}

// ----------------------------------------------------------------------------
// Comptage des alertes par cas — GET /v2/cases/{case_identifier}/alerts,
// avec parallélisme limité pour ne pas saturer l'API Mesh.
// ----------------------------------------------------------------------------
async function fetchAlertsCounts(baseUrl, token, caseIdentifiers) {
  const counts = new Map();
  let index = 0;

  async function worker() {
    while (index < caseIdentifiers.length) {
      const current = index++;
      const caseId = caseIdentifiers[current];
      try {
        const response = await fetch(`${baseUrl}${CASES_PATH}/${encodeURIComponent(caseId)}/alerts?page_size=1`, {
          headers: { Authorization: `Bearer ${token}` }
        });
        if (!response.ok) { counts.set(caseId, null); continue; }
        const data = await response.json();
        // On préfère un total explicite (pagination) si l'API le fournit, sinon
        // on retombe sur la longueur du tableau renvoyé.
        const total = data.total_count ?? data.total ?? data.count
          ?? (Array.isArray(data.data) ? data.data.length : (Array.isArray(data) ? data.length : null));
        counts.set(caseId, typeof total === 'number' ? total : null);
      } catch {
        counts.set(caseId, null);
      }
    }
  }

  const workers = Array.from({ length: Math.min(ALERTS_CONCURRENCY, caseIdentifiers.length) }, () => worker());
  await Promise.all(workers);
  return counts;
}

// ----------------------------------------------------------------------------
// Normalisation d'un cas brut Mesh vers le format renvoyé au front-end.
// ----------------------------------------------------------------------------
function normalizeCase(raw, alertsCounts, skipAlerts, nowMs) {
  const caseId = raw.case_identifier || raw.id || null;
  const createdAt = raw.created_at || null;
  const delayDays = createdAt ? Math.max(0, Math.floor((nowMs - new Date(createdAt).getTime()) / 86400000)) : null;

  return {
    caseIdentifier: caseId,
    caseType: raw.case_type || null,
    caseTypeLabel: CASE_TYPE_LABELS[raw.case_type] || raw.case_type || 'Type inconnu',
    status: raw.case_stage?.display_name || null,
    decisionType: raw.case_stage?.decision_type ?? null, // POSITIVE / NEGATIVE / null
    customerRef: raw.customer?.external_identifier || null,
    assignee: extractAssignee(raw),
    alertsCount: skipAlerts ? null : (alertsCounts.get(caseId) ?? null),
    createdAt,
    updatedAt: raw.updated_at || null,
    delayDays
  };
}

// Le nom exact du champ "assigné à" n'est pas garanti par la documentation
// publique de GET /v2/cases (cf. À CONFIRMER en en-tête). On teste ici les
// noms/emplacements les plus plausibles ; si aucun ne correspond, le front
// affiche "Non assigné" plutôt que de planter. À corriger une fois le nom réel
// confirmé par un appel de test avec les vraies clés API.
function extractAssignee(raw) {
  const candidate =
    raw.assignee?.name ||
    raw.assignee?.display_name ||
    raw.assignee?.email ||
    (typeof raw.assignee === 'string' ? raw.assignee : null) ||
    raw.assigned_to?.name ||
    raw.assigned_to ||
    raw.owner?.name ||
    raw.owner ||
    null;
  return candidate || null; // le front affiche "Non assigné" si null
}

function jsonResponse(status, body) {
  return {
    status,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  };
}
