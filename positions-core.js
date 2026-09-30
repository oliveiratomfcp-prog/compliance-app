// =============================================
//  POSITIONS EN PORTEFEUILLE — logique pure
//  Compliance App — Eiffel Investment Group
// =============================================
// Aucune dépendance au DOM, à MSAL ni à Microsoft Graph : tout ce qui touche au
// réseau est injecté (objet `graph`), ce qui permet de tester ce fichier seul.
// Utilisé par app.js (Restricted List + dépôt des positions).

const PositionsCore = (() => {
  'use strict';

  const PLACEHOLDER = '—';
  const POSITIONS_LIST_NAME = 'positions portefeuilles';
  const FORBIDDEN_LIST_NAMES = Object.freeze([
    'NDA List',
    'Information privilégiée',
    'Historique Compliance',
    'Registre cadeaux'
  ]);
  const MAX_CELL_LENGTH = 255;
  const BATCH_SIZE = 20;
  const MAX_ATTEMPTS = 6;
  const ALLOWED_EXTENSIONS = Object.freeze(['xlsx', 'xls']);
  const POSITION_TYPE = 'En portefeuille';
  const POSITION_SOURCE = 'Positions';

  // -----------------------------------------------
  // OUTILS
  // -----------------------------------------------
  function toText(v) { return v === null || v === undefined ? '' : String(v); }
  function norm(v) { return toText(v).trim().toLowerCase(); }
  function isEmptyValue(v) { const n = toText(v).trim(); return n === '' || n === PLACEHOLDER; }

  function escapeHtml(v) {
    return toText(v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function fileExtension(fileName) {
    const m = /\.([^.]+)$/.exec(toText(fileName));
    return m ? m[1].toLowerCase() : '';
  }

  // -----------------------------------------------
  // PARSING DES FICHIERS
  // -----------------------------------------------

  // Signature binaire : un vrai classeur .xlsx est une archive ZIP ("PK"), un vrai
  // .xls un conteneur OLE2 (D0 CF 11 E0 A1 B1 1A E1). Tout autre contenu (export HTML ou
  // texte renommé en .xls, CSV...) est refusé : SheetJS devinerait le format à partir du
  // texte, avec un risque de lignes perdues sans erreur.
  function workbookSignature(bytes) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const starts = sig => sig.every((v, i) => b[i] === v);
    if (b.length >= 4 && starts([0x50, 0x4B, 0x03, 0x04])) return 'zip';
    if (b.length >= 8 && starts([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])) return 'ole2';
    return null;
  }

  // Texte d'une cellule SheetJS : valeur affichée (w) en priorité, sinon valeur brute.
  function cellText(cell) {
    if (!cell) return '';
    if (cell.w !== undefined && cell.w !== null) return String(cell.w).trim();
    if (cell.v !== undefined && cell.v !== null) return String(cell.v).trim();
    return '';
  }

  // Lit un classeur (.xlsx, .xls) : première feuille, ligne 1 ignorée (en-têtes),
  // colonne A = nom, B = ISIN, C = fonds, lues par position. Seules les lignes où
  // A, B et C sont toutes vides sont ignorées.
  function parsePositionsWorkbook(XLSX, fileName, arrayBuffer) {
    if (!XLSX || typeof XLSX.read !== 'function') throw new Error('Bibliothèque SheetJS non chargée.');
    const ext = fileExtension(fileName);
    if (!ALLOWED_EXTENSIONS.includes(ext)) {
      throw new Error(`Format non pris en charge : ${fileName} (formats acceptés : .xlsx, .xls).`);
    }
    const bytes = new Uint8Array(arrayBuffer);
    if (!workbookSignature(bytes)) {
      throw new Error(`${fileName} n'est pas un classeur Excel valide (fichier texte, HTML ou CSV renommé ?). Ouvrez-le dans Excel et enregistrez-le au format .xlsx.`);
    }
    const wb = XLSX.read(bytes, { type: 'array', cellText: true, cellDates: false, dense: false });

    const rows = [];
    const sheetName = wb.SheetNames && wb.SheetNames[0];
    const ws = sheetName ? wb.Sheets[sheetName] : null;
    if (ws) {
      // Dernière ligne calculée à partir des cellules réellement présentes, sans se fier à
      // la dimension déclarée dans le fichier (parfois fausse selon l'outil qui l'a produit).
      let lastRow = -1;
      Object.keys(ws).forEach(key => {
        if (key[0] === '!') return;
        const addr = XLSX.utils.decode_cell(key);
        if (addr.c <= 2 && addr.r > lastRow) lastRow = addr.r;
      });
      for (let r = 1; r <= lastRow; r++) {
        const nom = cellText(ws[XLSX.utils.encode_cell({ r, c: 0 })]);
        const isin = cellText(ws[XLSX.utils.encode_cell({ r, c: 1 })]);
        const fonds = cellText(ws[XLSX.utils.encode_cell({ r, c: 2 })]);
        if (nom === '' && isin === '' && fonds === '') continue;
        rows.push({ nom, isin, fonds, file: fileName, line: r + 1 });
      }
    }
    return { fileName, sheetName: sheetName || '', rows };
  }

  // Réunit plusieurs fichiers en une seule liste, dans l'ordre des fichiers.
  function mergeParsedFiles(results) {
    return (results || []).reduce((acc, res) => acc.concat(res.rows || []), []);
  }

  // -----------------------------------------------
  // CONTRÔLES ET APERÇU
  // -----------------------------------------------

  // Format ISIN (2 lettres + 9 alphanumériques + 1 chiffre) et clé de contrôle (Luhn).
  function isinLooksValid(value) {
    const s = toText(value);
    if (!/^[A-Z]{2}[A-Z0-9]{9}[0-9]$/.test(s)) return false;
    const digits = s.split('').map(ch => (/[A-Z]/.test(ch) ? String(ch.charCodeAt(0) - 55) : ch)).join('');
    let sum = 0;
    let double = false;
    for (let i = digits.length - 1; i >= 0; i--) {
      let d = digits.charCodeAt(i) - 48;
      if (double) { d *= 2; if (d > 9) d -= 9; }
      sum += d;
      double = !double;
    }
    return sum % 10 === 0;
  }

  function positionKey(nom, isin, fonds) {
    return JSON.stringify([toText(nom).trim(), toText(isin).trim(), toText(fonds).trim()]);
  }

  function countKeys(list) {
    const m = new Map();
    list.forEach(p => {
      const k = positionKey(p.nom, p.isin, p.fonds);
      m.set(k, (m.get(k) || 0) + 1);
    });
    return m;
  }

  function multisetDiff(a, b) {
    const out = [];
    a.forEach((count, key) => {
      const diff = count - (b.get(key) || 0);
      if (diff > 0) {
        const [nom, isin, fonds] = JSON.parse(key);
        out.push({ nom, isin, fonds, count: diff });
      }
    });
    return out;
  }

  // currentPositions : [{nom, isin, fonds}] (valeurs SharePoint actuelles)
  function buildPreview(parsedFiles, currentPositions) {
    const files = parsedFiles || [];
    const rows = mergeParsedFiles(files);
    const byFile = files.map(f => ({ fileName: f.fileName, count: (f.rows || []).length }));

    const fundMap = new Map();
    rows.forEach(r => fundMap.set(r.fonds, (fundMap.get(r.fonds) || 0) + 1));
    const byFund = [...fundMap.entries()]
      .map(([fonds, count]) => ({ fonds, count }))
      .sort((x, y) => x.fonds.localeCompare(y.fonds, 'fr'));

    const warnings = { isinSuspect: [], emptyName: [], emptyFund: [], duplicates: [] };
    rows.forEach(r => {
      if (r.isin !== '' && !isinLooksValid(r.isin)) warnings.isinSuspect.push(r);
      if (r.nom === '') warnings.emptyName.push(r);
      if (r.fonds === '') warnings.emptyFund.push(r);
    });
    const dupGroups = new Map();
    rows.forEach(r => {
      const k = positionKey(r.nom, r.isin, r.fonds);
      if (!dupGroups.has(k)) dupGroups.set(k, []);
      dupGroups.get(k).push(r);
    });
    dupGroups.forEach(group => { if (group.length > 1) warnings.duplicates.push(group); });

    const blocking = { empty: rows.length === 0, tooLong: [] };
    rows.forEach(r => {
      ['nom', 'isin', 'fonds'].forEach(field => {
        if (r[field].length > MAX_CELL_LENGTH) blocking.tooLong.push({ row: r, field, length: r[field].length });
      });
    });

    const current = (currentPositions || []).map(p => ({ nom: toText(p.nom).trim(), isin: toText(p.isin).trim(), fonds: toText(p.fonds).trim() }));
    const newCounts = countKeys(rows);
    const curCounts = countKeys(current);

    return {
      rows,
      total: rows.length,
      byFile,
      byFund,
      warnings,
      blocking,
      isBlocked: blocking.empty || blocking.tooLong.length > 0,
      added: multisetDiff(newCounts, curCounts),
      removed: multisetDiff(curCounts, newCounts),
      currentTotal: current.length
    };
  }

  // -----------------------------------------------
  // GARDE DE SUPPRESSION
  // -----------------------------------------------
  // Seules les cibles créées par makePositionsTarget() sont acceptées par la garde :
  // un objet fabriqué à la main (même avec les bons identifiants) est refusé.
  const issuedTargets = new WeakSet();

  class GuardError extends Error {
    constructor(message) { super(message); this.name = 'GuardError'; }
  }

  function makePositionsTarget({ siteId, listId, displayName, forbiddenSiteIds, forbiddenListIds }) {
    if (typeof siteId !== 'string' || siteId === '') throw new GuardError('Site cible invalide.');
    if (typeof listId !== 'string' || listId === '') throw new GuardError('Liste cible invalide.');
    if (displayName !== POSITIONS_LIST_NAME) throw new GuardError(`Liste cible refusée : "${displayName}".`);
    if (FORBIDDEN_LIST_NAMES.includes(displayName)) throw new GuardError(`Liste protégée : "${displayName}".`);
    const fSites = Object.freeze([...(forbiddenSiteIds || [])].filter(Boolean));
    const fLists = Object.freeze([...(forbiddenListIds || [])].filter(Boolean));
    if (fSites.includes(siteId)) throw new GuardError('Site cible protégé (NDA / information privilégiée).');
    if (fLists.includes(listId)) throw new GuardError('Liste cible protégée.');
    const target = Object.freeze({ siteId, listId, displayName, forbiddenSiteIds: fSites, forbiddenListIds: fLists });
    issuedTargets.add(target);
    return target;
  }

  // Fonction unique de contrôle avant toute suppression : lève une GuardError si la
  // suppression ne vise pas exactement la liste "positions portefeuilles" résolue au démarrage.
  function assertDeletable(target, siteId, listId, itemId) {
    if (!target || !issuedTargets.has(target)) throw new GuardError('Suppression refusée : cible non résolue.');
    if (target.displayName !== POSITIONS_LIST_NAME) throw new GuardError('Suppression refusée : liste cible inattendue.');
    if (siteId !== target.siteId) throw new GuardError('Suppression refusée : site différent de CPLDashboard.');
    if (listId !== target.listId) throw new GuardError('Suppression refusée : liste différente de "positions portefeuilles".');
    if (target.forbiddenSiteIds.includes(siteId)) throw new GuardError('Suppression refusée : site protégé.');
    if (target.forbiddenListIds.includes(listId)) throw new GuardError('Suppression refusée : liste protégée.');
    if (!/^[1-9][0-9]*$/.test(toText(itemId))) throw new GuardError(`Suppression refusée : identifiant d'élément invalide (${toText(itemId)}).`);
    return true;
  }

  function itemsUrl(target) {
    return `/sites/${target.siteId}/lists/${target.listId}/items`;
  }

  function buildDeleteRequest(target, siteId, listId, itemId, requestId) {
    assertDeletable(target, siteId, listId, itemId);
    return { id: String(requestId), method: 'DELETE', url: `/sites/${siteId}/lists/${listId}/items/${itemId}` };
  }

  function buildCreateRequest(target, row, idDepot, requestId) {
    if (!target || !issuedTargets.has(target)) throw new GuardError('Création refusée : cible non résolue.');
    return {
      id: String(requestId),
      method: 'POST',
      url: itemsUrl(target),
      headers: { 'Content-Type': 'application/json' },
      body: { fields: { Title: row.nom, ISIN: row.isin, Fonds: row.fonds, IdDepot: idDepot } }
    };
  }

  // Second verrou, appliqué à chaque sous-requête juste avant l'envoi du $batch.
  function assertBatchRequestAllowed(req, target) {
    if (!target || !issuedTargets.has(target)) throw new GuardError('Requête refusée : cible non résolue.');
    const base = itemsUrl(target);
    const url = toText(req && req.url);
    const method = toText(req && req.method).toUpperCase();
    if (method === 'POST') {
      if (url !== base) throw new GuardError(`Requête refusée : création hors de la liste des positions (${url}).`);
      return true;
    }
    if (method === 'DELETE') {
      const prefix = base + '/';
      if (!url.startsWith(prefix) || !/^[1-9][0-9]*$/.test(url.slice(prefix.length))) {
        throw new GuardError(`Requête refusée : suppression hors de la liste des positions (${url}).`);
      }
      return true;
    }
    throw new GuardError(`Requête refusée : méthode non autorisée (${method}).`);
  }

  // -----------------------------------------------
  // ENVOI $batch AVEC NOUVELLES TENTATIVES (429 / 503)
  // -----------------------------------------------
  function parseRetryAfter(value) {
    if (value === undefined || value === null || value === '') return null;
    const n = Number(value);
    if (Number.isFinite(n) && n >= 0) return n * 1000;
    const d = Date.parse(value);
    if (!Number.isNaN(d)) return Math.max(0, d - Date.now());
    return null;
  }

  function headerValue(headers, name) {
    if (!headers) return undefined;
    const key = Object.keys(headers).find(k => k.toLowerCase() === name.toLowerCase());
    return key ? headers[key] : undefined;
  }

  function backoffMs(attempt) { return Math.min(60000, 1000 * Math.pow(2, attempt)); }
  function isRetryable(status) { return status === 429 || status === 503; }

  // graph.postBatch({requests}) -> {status, retryAfter, json}
  // Renvoie une Map id -> {status, body}. Seules les réponses 429 / 503 sont retentées :
  // tout autre échec est remonté tel quel, pour ne jamais créer de doublon fantôme.
  async function executeBatch(graph, target, requests, opts = {}) {
    const sleep = opts.sleep || (ms => new Promise(res => setTimeout(res, ms)));
    const maxAttempts = opts.maxAttempts || MAX_ATTEMPTS;
    if (requests.length > BATCH_SIZE) throw new Error(`Paquet trop grand (${requests.length} > ${BATCH_SIZE}).`);
    requests.forEach(r => assertBatchRequestAllowed(r, target));

    const results = new Map();
    let pending = requests.slice();
    for (let attempt = 1; attempt <= maxAttempts && pending.length; attempt++) {
      const res = await graph.postBatch({ requests: pending });
      if (isRetryable(res.status)) {
        if (attempt === maxAttempts) throw new Error(`Microsoft Graph indisponible (HTTP ${res.status}) après ${maxAttempts} tentatives.`);
        if (opts.onRetry) opts.onRetry({ attempt, status: res.status });
        await sleep(parseRetryAfter(res.retryAfter) ?? backoffMs(attempt));
        continue;
      }
      if (res.status < 200 || res.status >= 300) throw new Error(`Erreur $batch (HTTP ${res.status}).`);

      const byId = new Map(((res.json && res.json.responses) || []).map(r => [String(r.id), r]));
      const retry = [];
      let waitMs = 0;
      pending.forEach(req => {
        const sub = byId.get(req.id);
        if (!sub) {
          results.set(req.id, { status: 0, body: { error: { message: 'Réponse manquante dans le $batch.' } } });
        } else if (isRetryable(sub.status) && attempt < maxAttempts) {
          retry.push(req);
          const ra = parseRetryAfter(headerValue(sub.headers, 'Retry-After'));
          waitMs = Math.max(waitMs, ra ?? backoffMs(attempt));
        } else {
          results.set(req.id, { status: sub.status, body: sub.body });
        }
      });
      pending = retry;
      if (pending.length) {
        if (opts.onRetry) opts.onRetry({ attempt, status: 429, count: pending.length });
        await sleep(waitMs);
      }
    }
    return results;
  }

  function chunk(list, size) {
    const out = [];
    for (let i = 0; i < list.length; i += size) out.push(list.slice(i, i + size));
    return out;
  }

  // Supprime des éléments (via la garde). Une réponse 404 compte comme supprimée
  // (l'élément n'existe déjà plus).
  async function deleteItems(graph, target, itemIds, opts = {}) {
    const requests = itemIds.map((id, i) => buildDeleteRequest(target, target.siteId, target.listId, id, i + 1));
    let deleted = 0;
    const failed = [];
    for (const part of chunk(requests, BATCH_SIZE)) {
      let results;
      try {
        results = await executeBatch(graph, target, part, opts);
      } catch (e) {
        if (e instanceof GuardError) throw e;
        part.forEach(r => failed.push({ url: r.url, status: 0, message: e.message }));
        continue;
      }
      part.forEach(r => {
        const res = results.get(r.id);
        if (res && ((res.status >= 200 && res.status < 300) || res.status === 404)) deleted++;
        else failed.push({ url: r.url, status: res ? res.status : 0 });
      });
      if (opts.onProgress) opts.onProgress({ phase: opts.phase || 'delete', done: deleted + failed.length, total: requests.length });
    }
    return { deleted, failed };
  }

  function generateIdDepot(cryptoObj, now) {
    const c = cryptoObj || (typeof crypto !== 'undefined' ? crypto : null);
    let rand;
    if (c && typeof c.randomUUID === 'function') {
      rand = c.randomUUID();
    } else if (c && typeof c.getRandomValues === 'function') {
      rand = Array.from(c.getRandomValues(new Uint8Array(16)), b => b.toString(16).padStart(2, '0')).join('');
    } else {
      throw new Error('Générateur aléatoire indisponible.');
    }
    return `${(now || new Date()).toISOString()}_${rand}`;
  }

  // -----------------------------------------------
  // REMPLACEMENT SÉCURISÉ
  // -----------------------------------------------
  // graph : { postBatch({requests}), listItems(target) -> [{id, fields}] }
  // 1. création de toutes les lignes avec le nouvel IdDepot
  // 2. vérification : exactement rows.length lignes portent le nouvel IdDepot
  // 3. seulement alors : suppression des lignes dont l'IdDepot est différent
  // En cas d'échec en 1 ou 2 : suppression des lignes du nouvel IdDepot, ancienne liste intacte.
  async function runReplacement({ graph, target, rows, idDepot, onProgress, sleep, maxAttempts }) {
    if (!target || !issuedTargets.has(target)) throw new GuardError('Remplacement refusé : cible non résolue.');
    if (!Array.isArray(rows) || rows.length === 0) throw new Error('Remplacement refusé : aucune ligne à déposer.');
    const tooLong = rows.some(r => ['nom', 'isin', 'fonds'].some(f => toText(r[f]).length > MAX_CELL_LENGTH));
    if (tooLong) throw new Error(`Remplacement refusé : cellule de plus de ${MAX_CELL_LENGTH} caractères.`);
    if (typeof idDepot !== 'string' || idDepot === '') throw new Error('IdDepot invalide.');

    const opts = { sleep, maxAttempts };
    const progress = p => { if (onProgress) onProgress(p); };
    const createdIds = [];
    let failure = null;

    // Étape 1 : création
    const requests = rows.map((row, i) => buildCreateRequest(target, row, idDepot, i + 1));
    let done = 0;
    for (const part of chunk(requests, BATCH_SIZE)) {
      let results;
      try {
        results = await executeBatch(graph, target, part, opts);
      } catch (e) {
        if (e instanceof GuardError) throw e;
        failure = { step: 'create', message: e.message };
        break;
      }
      part.forEach(r => {
        const res = results.get(r.id);
        if (res && res.status === 201 && res.body && res.body.id) createdIds.push(String(res.body.id));
        else if (!failure) failure = { step: 'create', message: `Création refusée pour la ligne ${r.id} (HTTP ${res ? res.status : '?'})${res && res.body && res.body.error ? ' : ' + res.body.error.message : ''}.` };
      });
      done += part.length;
      progress({ phase: 'create', done, total: requests.length });
      if (failure) break;
    }

    // Étape 2 : vérification par relecture
    let items = null;
    if (!failure) {
      try {
        items = await graph.listItems(target);
        const countNew = items.filter(it => (it.fields || {}).IdDepot === idDepot).length;
        if (countNew !== rows.length) {
          failure = { step: 'verify', message: `Vérification échouée : ${countNew} ligne(s) trouvée(s) pour ${rows.length} attendue(s).` };
        }
      } catch (e) {
        failure = { step: 'verify', message: `Relecture impossible : ${e.message}` };
      }
    }

    if (failure) {
      // Retour arrière : suppression des lignes partielles du nouvel IdDepot
      let toDelete = createdIds.slice();
      let rollbackReadFailed = false;
      try {
        const current = await graph.listItems(target);
        current.filter(it => (it.fields || {}).IdDepot === idDepot).forEach(it => {
          if (!toDelete.includes(String(it.id))) toDelete.push(String(it.id));
        });
      } catch (e) {
        rollbackReadFailed = true;
      }
      const rb = await deleteItems(graph, target, toDelete, { ...opts, onProgress, phase: 'rollback' });
      return {
        ok: false,
        failure,
        created: createdIds.length,
        rolledBack: rb.deleted,
        rollbackFailed: rb.failed.length,
        rollbackReadFailed,
        idDepot
      };
    }

    // Étape 3 : suppression de l'ancienne liste (IdDepot différent du nouveau)
    const oldIds = items.filter(it => (it.fields || {}).IdDepot !== idDepot).map(it => String(it.id));
    const del = await deleteItems(graph, target, oldIds, { ...opts, onProgress, phase: 'delete' });
    return {
      ok: true,
      created: createdIds.length,
      deleted: del.deleted,
      deleteFailed: del.failed.length,
      toDelete: oldIds.length,
      idDepot
    };
  }

  // -----------------------------------------------
  // RESTRICTED LIST : correspondances et affichage
  // -----------------------------------------------

  // Élément SharePoint -> ligne de la Restricted List
  function positionToRestrictedItem(item) {
    const f = (item && item.fields) || {};
    const val = v => (isEmptyValue(v) ? PLACEHOLDER : toText(v).trim());
    return {
      nom: val(f.Title),
      isin: val(f.ISIN),
      fonds: val(f.Fonds),
      dateDebut: PLACEHOLDER,
      dateFin: PLACEHOLDER,
      equipe: PLACEHOLDER,
      signataire: PLACEHOLDER,
      type: POSITION_TYPE,
      source: POSITION_SOURCE
    };
  }

  // Correspondance sur le nom du titre et l'ISIN uniquement (vérification en masse,
  // attestation, demande de transaction). Une valeur vide ne correspond jamais.
  function matchesSecurity(item, query) {
    const q = norm(query);
    if (!q) return false;
    const nom = isEmptyValue(item.nom) ? '' : norm(item.nom);
    const isin = isEmptyValue(item.isin) ? '' : norm(item.isin);
    return (nom !== '' && nom.includes(q)) || (isin !== '' && isin.includes(q));
  }

  // Vérification en masse : correspondance du titre, ou ISIN contenu dans la saisie.
  function matchesBulkEntry(item, entry) {
    if (matchesSecurity(item, entry)) return true;
    const e = norm(entry);
    const isin = isEmptyValue(item.isin) ? '' : norm(item.isin);
    return e !== '' && isin !== '' && e.includes(isin);
  }

  // Recherche du tableau et export : nom, ISIN et nom du fonds.
  function matchesSearch(item, query) {
    const q = norm(query);
    if (!q) return true;
    if (matchesSecurity(item, q)) return true;
    const fonds = isEmptyValue(item.fonds) ? '' : norm(item.fonds);
    return fonds !== '' && fonds.includes(q);
  }

  function isPositionItem(item) { return !!item && item.source === POSITION_SOURCE; }

  // Libellé des types d'une liste de correspondances, sans doublon :
  // "NDA, En portefeuille (Fonds A, Fonds B)"
  function describeMatchTypes(matches) {
    const types = [];
    const funds = [];
    (matches || []).forEach(m => {
      if (isPositionItem(m)) {
        if (!types.includes(POSITION_TYPE)) types.push(POSITION_TYPE);
        const f = isEmptyValue(m.fonds) ? '' : toText(m.fonds).trim();
        if (f && !funds.includes(f)) funds.push(f);
      } else if (m && m.type && !types.includes(m.type)) {
        types.push(m.type);
      }
    });
    return types.map(tp => (tp === POSITION_TYPE && funds.length ? `${tp} (${funds.join(', ')})` : tp)).join(', ');
  }

  // Nombre de dépôts distincts (IdDepot) présents dans la liste
  function distinctIdDepots(items) {
    const set = new Set();
    (items || []).forEach(it => set.add(toText(((it && it.fields) || {}).IdDepot)));
    return [...set];
  }

  return Object.freeze({
    PLACEHOLDER, POSITIONS_LIST_NAME, FORBIDDEN_LIST_NAMES, MAX_CELL_LENGTH, BATCH_SIZE,
    ALLOWED_EXTENSIONS, POSITION_TYPE, POSITION_SOURCE, GuardError,
    escapeHtml, fileExtension, isEmptyValue,
    workbookSignature, parsePositionsWorkbook, mergeParsedFiles,
    isinLooksValid, buildPreview,
    makePositionsTarget, assertDeletable, buildDeleteRequest, buildCreateRequest, assertBatchRequestAllowed,
    parseRetryAfter, executeBatch, deleteItems, generateIdDepot, runReplacement,
    positionToRestrictedItem, matchesSecurity, matchesBulkEntry, matchesSearch, isPositionItem,
    describeMatchTypes, distinctIdDepots
  });
})();
