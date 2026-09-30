// Authentification des Azure Functions.
// Chaque appel doit porter le jeton Microsoft Graph de l'utilisateur connecté
// (en-tête X-EIG-Graph-Token, ou Authorization: Bearer). La Function appelle
// /me avec ce jeton : l'appel doit réussir et le userPrincipalName doit se
// terminer par @eiffel-ig.com (les comptes invités sont donc exclus).
// Le résultat est mis en cache quelques minutes, indexé par l'empreinte du
// jeton (le jeton lui-même n'est jamais conservé ni journalisé).

const fetch = require('node-fetch');
const crypto = require('crypto');

const ALLOWED_DOMAIN = '@eiffel-ig.com';
const CACHE_TTL_MS = 5 * 60 * 1000;
const CACHE_MAX_ENTRIES = 500;
const MAX_TOKEN_LENGTH = 12000;

const cache = new Map();

function headerValue(req, name) {
  const headers = (req && req.headers) || {};
  const key = Object.keys(headers).find(function (k) { return k.toLowerCase() === name; });
  return key ? String(headers[key] || '') : '';
}

function extractToken(req) {
  const custom = headerValue(req, 'x-eig-graph-token').trim();
  if (custom) return custom;
  const m = /^Bearer\s+(\S+)$/i.exec(headerValue(req, 'authorization').trim());
  return m ? m[1] : '';
}

// Date d'expiration du jeton (champ exp du JWT), sans vérification de signature :
// la validité réelle est établie par l'appel à /me.
function tokenExpiry(token) {
  try {
    const part = token.split('.')[1];
    if (!part) return null;
    const payload = JSON.parse(Buffer.from(part.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
    return typeof payload.exp === 'number' ? payload.exp * 1000 : null;
  } catch (e) {
    return null;
  }
}

function pruneCache(now) {
  cache.forEach(function (entry, key) { if (entry.expires <= now) cache.delete(key); });
  while (cache.size > CACHE_MAX_ENTRIES) cache.delete(cache.keys().next().value);
}

// Renvoie { email, name } ou null si l'utilisateur n'est pas authentifié.
async function authenticate(req) {
  const token = extractToken(req);
  if (!token || token.length > MAX_TOKEN_LENGTH) return null;

  const key = crypto.createHash('sha256').update(token).digest('hex');
  const now = Date.now();
  const hit = cache.get(key);
  if (hit && hit.expires > now) return hit.user;

  let me;
  try {
    const res = await fetch('https://graph.microsoft.com/v1.0/me?$select=userPrincipalName,displayName', {
      headers: { Authorization: 'Bearer ' + token },
      timeout: 10000
    });
    if (!res.ok) return null;
    me = await res.json();
  } catch (e) {
    return null;
  }

  const upn = String((me && me.userPrincipalName) || '').toLowerCase();
  if (!upn.endsWith(ALLOWED_DOMAIN) || upn.indexOf('#ext#') !== -1) return null;

  const user = { email: upn, name: String(me.displayName || upn) };
  const exp = tokenExpiry(token);
  cache.set(key, { user: user, expires: Math.min(now + CACHE_TTL_MS, exp || now + CACHE_TTL_MS) });
  pruneCache(now);
  return user;
}

function sendJson(context, status, body) {
  context.res = {
    status: status,
    headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
    body: JSON.stringify(body)
  };
}

function unauthorized(context) {
  sendJson(context, 401, { error: { message: 'Authentification requise : reconnectez-vous au Compliance Portal.' } });
}

module.exports = { authenticate: authenticate, sendJson: sendJson, unauthorized: unauthorized, extractToken: extractToken };
