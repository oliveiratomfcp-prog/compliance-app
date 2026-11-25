// =============================================
//  TRACKER CPL : logique pure de la recherche IA et des demandes similaires
//  Compliance App, Eiffel Investment Group
// =============================================
// Aucune dépendance au DOM, à MSAL ni à Graph. Construit le catalogue envoyé à l'IA à partir
// des demandes déjà chargées et des commentaires, en respectant la visibilité de l'utilisateur :
//   - Compliance : toutes les demandes ;
//   - collaborateur : ses demandes et celles où il est en copie ; pour les demandes similaires
//     seulement, les demandes d'autres personnes marquées "Partageable" par la Compliance, sans
//     nom ni département du demandeur.
// Les notes internes (préfixe [INTERNE]) n'entrent jamais dans le catalogue. La "réponse
// apportée" est faite des commentaires publics des membres de la Compliance et, pour une
// demande rejetée, du motif de rejet : c'est ce texte, tel quel, que le navigateur affiche.

const TrackerAiCore = (() => {
  'use strict';

  const INTERNE = '[INTERNE]';
  const REPLY_MAX = 3000;       // réponse affichée à l'utilisateur
  const REPLY_EXCERPT = 450;    // extrait envoyé à l'IA (recherche)
  const REPLY_EXCERPT_SIMILAR = 750;   // extrait envoyé à l'IA (demandes similaires : synthèse)
  const MSG_EXCERPT = 300;      // longueur maximale d'un message dans l'extrait
  const DESC_EXCERPT = 350;

  function truncate(s, n) {
    s = String(s || '');
    return s.length > n ? s.slice(0, n - 1).trimEnd() + '…' : s;
  }

  // HTML (éditeur de description ou de commentaire) -> texte brut
  function htmlToText(html) {
    return String(html || '')
      .replace(/<\s*(br|\/p|\/div|\/li|\/h[1-6])\s*\/?>/gi, '\n')
      .replace(/<li[^>]*>/gi, '- ')
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/gi, ' ').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
      .replace(/&quot;/gi, '"').replace(/&#39;/g, "'").replace(/&amp;/gi, '&')
      .replace(/[ \t ]+/g, ' ')
      .replace(/ *\n */g, '\n').replace(/\n{3,}/g, '\n\n')
      .trim();
  }

  // Commentaire stocké (HTML + éventuel lien markdown de pièce jointe) -> texte lisible
  function cleanComment(raw) {
    return htmlToText(raw)
      .replace(/📎\s*\[([^\]]+)\]\([^)\s]*\)/g, '[pièce jointe : $1]')
      .replace(/📎\s*([^\n]+)/g, '[pièce jointe : $1]')
      .trim();
  }

  function isInternal(raw) { return String(raw || '').toUpperCase().indexOf(INTERNE) !== -1; }

  function isCplAuthor(item, cplEmails, cplNames) {
    const u = (item && item.createdBy && item.createdBy.user) || {};
    const mail = String(u.email || u.userPrincipalName || '').toLowerCase();
    if (mail && (cplEmails || []).some(e => String(e).toLowerCase() === mail)) return true;
    const f = (item && item.fields) || {};
    const names = [u.displayName, f.NomAuteur].map(n => String(n || '').trim().toLowerCase()).filter(Boolean);
    return names.some(n => (cplNames || []).some(c => String(c).toLowerCase() === n));
  }

  // items : éléments bruts de la liste "Tracker CPL commentaires" (champs + createdBy).
  // Renvoie Map(référence -> [{ ts, text }]) des seuls commentaires publics de la Compliance.
  function buildCommentsIndex(items, opts) {
    const map = new Map();
    (items || []).forEach(item => {
      const f = item.fields || {};
      const ref = String(f.Title || '').trim();
      const raw = String(f.Texte || '');
      if (!ref || !raw || isInternal(raw)) return;
      if (!isCplAuthor(item, opts && opts.cplEmails, opts && opts.cplNames)) return;
      const text = cleanComment(raw);
      if (!text) return;
      const ts = item.createdDateTime ? Date.parse(item.createdDateTime) : 0;
      if (!map.has(ref)) map.set(ref, []);
      map.get(ref).push({ ts: isNaN(ts) ? 0 : ts, text });
    });
    map.forEach(list => list.sort((a, b) => a.ts - b.ts));
    return map;
  }

  function fmtDate(ts) {
    return ts ? new Date(ts).toLocaleDateString('fr-FR') : '';
  }

  // Messages de la Compliance dans l'ordre (commentaires publics puis, pour une demande rejetée,
  // le motif de rejet), tels qu'écrits. Leur numéro (à partir de 1) sert à l'IA pour désigner le
  // message décisif, que le navigateur affiche ensuite mot pour mot.
  function replyMessages(req, comments) {
    const out = (comments || []).map(c => (c.ts ? `[${fmtDate(c.ts)}] ` : '') + c.text);
    if (req && req.statut === 'Rejeté' && req.motifRejet) out.push('Motif de rejet : ' + req.motifRejet);
    return out;
  }

  // Réponse apportée par la Compliance, telle qu'écrite (jamais reformulée)
  function buildReply(req, comments) {
    return truncate(replyMessages(req, comments).join('\n\n'), REPLY_MAX);
  }

  // Extrait envoyé à l'IA : messages numérotés « #1 », « #2 »..., chacun tronqué. Si le total
  // dépasse le budget, on garde le premier message (le contexte) et les derniers (souvent la
  // réponse décisive), et on retire ceux du milieu.
  function replyExcerpt(messages, budget) {
    const items = messages.map((m, i) => `#${i + 1} ${truncate(m, MSG_EXCERPT)}`);
    const size = list => list.join(' ').length;
    if (size(items) <= budget) return items.join(' ');
    const kept = items.slice();
    while (kept.length > 2 && size(kept) > budget) kept.splice(1, 1);
    return truncate(kept.join(' '), budget);
  }

  const FINAL = ['Traité', 'Clôturé', 'Rejeté'];
  const SANS_OBJET = /annul|doublon|erreur|retir|cancel|duplicate|sans objet/i;

  // Demandes sans valeur pour un collaborateur qui prépare une nouvelle demande : annulées ou
  // rejetées pour doublon, ou terminées sans aucune réponse de la Compliance. Une demande encore
  // ouverte reste proposée, même sans réponse (risque de doublon).
  function sansInteret(req, messages) {
    if (!req) return true;
    if (req.statut === 'Rejeté' && SANS_OBJET.test(String(req.motifRejet || ''))) return true;
    return FINAL.indexOf(req.statut) !== -1 && !messages.length;
  }

  function followersOf(req) {
    return String((req && req.followers) || '').toLowerCase().split(',').map(e => e.trim()).filter(Boolean);
  }

  // La demande figure dans la liste de l'utilisateur (la sienne ou en copie)
  function visibleTo(req, user) {
    const email = String((user && user.email) || '').toLowerCase();
    if (!email || !req) return false;
    return String(req.email || '').toLowerCase() === email || followersOf(req).indexOf(email) !== -1;
  }

  function isShareable(req) { return !!req && req.partageable === true; }

  // mode "search" : demandes de la liste de l'utilisateur ; mode "similar" : en plus, pour un
  // collaborateur, les demandes partageables des autres (anonymisées).
  // Renvoie { entries, byRef } : entries part à l'IA, byRef sert à afficher les résultats.
  function buildCatalogue(requests, commentsIndex, opts) {
    const entries = [];
    const byRef = new Map();
    (requests || []).forEach(r => {
      if (!r || !r.ref) return;
      const own = !!opts.isCPL || visibleTo(r, opts);
      const shared = !own && opts.mode === 'similar' && isShareable(r);
      if (!own && !shared) return;
      const comments = commentsIndex && commentsIndex.get(r.ref);
      const messages = replyMessages(r, comments);
      if (opts.mode === 'similar' && sansInteret(r, messages)) return;
      const reply = buildReply(r, comments);
      const e = {
        ref: r.ref, objet: r.objet, type: r.type, statut: r.statut,
        description: truncate(htmlToText(r.desc), DESC_EXCERPT),
        reponse: opts.mode === 'similar' ? replyExcerpt(messages, REPLY_EXCERPT_SIMILAR) : truncate(reply, REPLY_EXCERPT)
      };
      if (!shared) { e.dept = r.dept; e.demandeur = r.from; }
      entries.push(e);
      byRef.set(r.ref, { req: r, shared, reply, messages });
    });
    return { entries, byRef };
  }

  // Empreinte du contenu du catalogue (clé du cache de résultats côté navigateur)
  function catalogueSignature(entries) {
    let h = 5381;
    const s = entries.map(e => `${e.ref}|${e.statut}|${e.objet.length}|${e.description.length}|${e.reponse.length}`).join(';');
    for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
    return `${entries.length}:${(h >>> 0).toString(36)}`;
  }

  function parseSimilaires(raw) {
    try {
      const o = JSON.parse(raw);
      return o && Array.isArray(o.propose) ? o : null;
    } catch (e) { return null; }
  }

  return Object.freeze({
    htmlToText, cleanComment, isInternal, isCplAuthor, buildCommentsIndex, buildReply, replyMessages, replyExcerpt, sansInteret, visibleTo, isShareable,
    buildCatalogue, catalogueSignature, parseSimilaires, truncate, REPLY_MAX
  });
})();
