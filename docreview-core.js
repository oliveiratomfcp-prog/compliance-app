// =============================================
//  DOC REVIEW : logique pure (navigateur)
//  Compliance App — Eiffel Investment Group
// =============================================
// Aucune dépendance au DOM, à MSAL ni à Microsoft Graph : contrôle des citations,
// structure PPTX, correspondance pages / slides, mise en forme de l'analyse,
// rapport PDF. Testable seul.

const DocReviewCore = (() => {
  'use strict';

  // Appréciation globale de l'IA et valeur enregistrée dans la colonne Score (3, 2 ou 1).
  // Libellés identiques au schéma serveur (api/shared/docreview-logic.js).
  const APPRECIATIONS = Object.freeze(['Prêt à soumettre', 'Quelques ajustements conseillés', 'À retravailler']);
  const SCORE_PAR_APPRECIATION = Object.freeze({ 'Prêt à soumettre': 3, 'Quelques ajustements conseillés': 2, 'À retravailler': 1 });
  const NIVEAUX = Object.freeze(['À traiter', 'Recommandé', 'Suggestion']);
  const CONFIANCES = ['élevé', 'moyen', 'faible'];

  // Paramètres d'analyse
  const ANALYSE_CONFIG = Object.freeze({
    pagesParLot: 6,          // pages envoyées ensemble à l'IA (étape 1)
    lotsEnParallele: 2,      // lots traités en même temps
    grandCotePx: 1800,       // résolution des images de page envoyées à l'IA
    qualiteJpeg: 0.85,
    vignettePx: 320,         // vignettes affichées dans le rapport
    qualiteVignette: 0.7,
    seuilPagesAvertissement: 80,
    maxTexteParPage: 20000,  // doit rester égal à la limite serveur (api/shared/docreview-logic.js)
    intervalleSuiviMs: 3000,
    delaiMaxEtapeMs: 15 * 60 * 1000
  });

  function toText(v) { return v === null || v === undefined ? '' : String(v); }

  function escapeHtml(v) {
    return toText(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function scoreFromAppreciation(valeur) {
    return SCORE_PAR_APPRECIATION[valeur] || null;
  }

  // Avertissement non bloquant : fonds en principe réservé aux professionnels et public Retail
  function isProfessionalFundForRetail(typeFonds, marche, fondsProfessionnels) {
    return marche === 'retail' && (fondsProfessionnels || []).indexOf(typeFonds) !== -1;
  }

  function makeRef(prefix, date, spItemId) {
    return `${prefix}-${date.getFullYear()}-${String(spItemId).padStart(4, '0')}`;
  }

  // Nom de fichier accepté par SharePoint (caractères interdits remplacés, longueur bornée)
  function sanitizeFileName(name) {
    let n = toText(name).normalize('NFC').replace(/["*:<>?/\\|\u0000-\u001f]/g, '_').replace(/\s+/g, ' ').trim();
    n = n.replace(/^[.\s~$]+/, '').replace(/[.\s]+$/, '');
    if (!n) n = 'document';
    if (n.length > 150) {
      const dot = n.lastIndexOf('.');
      const ext = dot > 0 && n.length - dot <= 10 ? n.slice(dot) : '';
      n = n.slice(0, 150 - ext.length) + ext;
    }
    return n;
  }

  function extensionOf(name) {
    const m = /\.([^.]+)$/.exec(toText(name));
    return m ? m[1].toLowerCase() : '';
  }

  function baseName(name) {
    const n = toText(name);
    const dot = n.lastIndexOf('.');
    return dot > 0 ? n.slice(0, dot) : n;
  }

  // ---------------------------------------------------------------------------
  // Contrôle des citations : une citation "mot pour mot" doit se retrouver dans le
  // texte extrait de la page (espaces, apostrophes et guillemets normalisés).
  // ---------------------------------------------------------------------------
  function normalizeForMatch(s) {
    return toText(s).normalize('NFKC').toLowerCase()
      .replace(/[‘’‚‛′`´]/g, "'")
      .replace(/[“”„‟«»″]/g, '"')
      .replace(/[‐-―−]/g, '-')
      .replace(/…/g, '...')
      .replace(/["']/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // true : retrouvée ; false : absente d'une page qui contient du texte ; null : non vérifiable
  function checkCitation(citation, pages, pageTexts) {
    const c = normalizeForMatch(citation);
    if (!c) return null;
    const candidates = (pages && pages.length ? pages : Object.keys(pageTexts).map(Number));
    let verifiable = false;
    for (const p of candidates) {
      const t = normalizeForMatch(pageTexts[p]);
      if (!t) continue;
      verifiable = true;
      if (t.indexOf(c) !== -1) return true;
    }
    return verifiable ? false : null;
  }

  // ---------------------------------------------------------------------------
  // Lots de pages
  // ---------------------------------------------------------------------------
  function buildBatches(totalPages, size) {
    const out = [];
    for (let start = 1; start <= totalPages; start += size) {
      const pages = [];
      for (let p = start; p < start + size && p <= totalPages; p++) pages.push(p);
      out.push(pages);
    }
    return out;
  }

  // ---------------------------------------------------------------------------
  // PPTX : ordre réel des slides (presentation.xml) et slides masquées (show="0")
  // ---------------------------------------------------------------------------
  const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';

  function resolveTarget(baseDir, target) {
    const parts = (target.startsWith('/') ? target.slice(1) : baseDir + '/' + target).split('/');
    const out = [];
    parts.forEach(p => { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
    return out.join('/');
  }

  async function parsePptxStructure(JSZip, arrayBuffer) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const presFile = zip.file('ppt/presentation.xml');
    const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
    if (!presFile || !relsFile) throw new Error('Structure PPTX invalide (presentation.xml introuvable).');
    const parser = new DOMParser();
    const pres = parser.parseFromString(await presFile.async('string'), 'application/xml');
    const rels = parser.parseFromString(await relsFile.async('string'), 'application/xml');
    const targets = {};
    Array.from(rels.getElementsByTagNameNS(NS_REL, 'Relationship')).forEach(r => { targets[r.getAttribute('Id')] = r.getAttribute('Target'); });
    const ids = Array.from(pres.getElementsByTagNameNS(NS_P, 'sldId'));
    const slides = [];
    for (let i = 0; i < ids.length; i++) {
      const rid = ids[i].getAttributeNS(NS_R, 'id');
      const target = targets[rid];
      if (!target) throw new Error('Structure PPTX invalide (relation de slide introuvable).');
      const path = resolveTarget('ppt', target);
      const f = zip.file(path);
      if (!f) throw new Error('Structure PPTX invalide (slide introuvable : ' + path + ').');
      const sld = parser.parseFromString(await f.async('string'), 'application/xml');
      const root = sld.documentElement;
      const hidden = root && (root.getAttribute('show') === '0' || root.getAttribute('show') === 'false');
      slides.push({ position: i + 1, path, hidden: !!hidden });
    }
    const size = pres.getElementsByTagNameNS(NS_P, 'sldSz')[0];
    return {
      slides,
      total: slides.length,
      hidden: slides.filter(s => s.hidden).map(s => s.position),
      widthEmu: size ? Number(size.getAttribute('cx')) : null,
      heightEmu: size ? Number(size.getAttribute('cy')) : null
    };
  }

  // Correspondance pages du PDF converti / slides du PPTX. La conversion peut inclure
  // ou exclure les slides masquées : on déduit la règle du nombre de pages obtenu.
  function mapPdfPages(pdfPageCount, sourceType, pptx) {
    const labels = [];
    if (sourceType === 'pptx' && pptx) {
      const visibles = pptx.slides.filter(s => !s.hidden).map(s => s.position);
      if (pdfPageCount === pptx.total) {
        for (let i = 1; i <= pdfPageCount; i++) labels.push({ page: i, libelle: 'Slide ' + i, slide: i });
        return { mode: 'toutes', incertaine: false, labels };
      }
      if (pdfPageCount === visibles.length) {
        visibles.forEach((s, i) => labels.push({ page: i + 1, libelle: 'Slide ' + s, slide: s }));
        return { mode: 'visibles', incertaine: false, labels };
      }
      for (let i = 1; i <= pdfPageCount; i++) labels.push({ page: i, libelle: 'Page ' + i, slide: null });
      return { mode: 'incertaine', incertaine: true, labels };
    }
    for (let i = 1; i <= pdfPageCount; i++) labels.push({ page: i, libelle: 'Page ' + i, slide: null });
    return { mode: 'pages', incertaine: false, labels };
  }

  // ---------------------------------------------------------------------------
  // Consolidation : remarques hiérarchisées, contrôle des pages et des citations,
  // appréciation globale et compteurs enregistrés dans SharePoint
  // ---------------------------------------------------------------------------
  function cleanPages(list, totalPages) {
    const out = [];
    (Array.isArray(list) ? list : []).forEach(p => {
      if (Number.isInteger(p) && p >= 1 && p <= totalPages && out.indexOf(p) === -1) out.push(p);
    });
    return out.sort((a, b) => a - b);
  }

  function nullableText(v) {
    const t = toText(v).trim();
    return t ? t : null;
  }

  function textList(list) {
    return (Array.isArray(list) ? list : []).map(v => toText(v).trim()).filter(Boolean);
  }

  function consolidate(analyse, pageTexts, totalPages) {
    const a = analyse || {};
    const comp = a.comprehension || {};
    const app = a.appreciation_globale || {};
    const remarques = (Array.isArray(a.remarques) ? a.remarques : []).map((r, i) => {
      const pages = cleanPages(r.pages, totalPages);
      const element = nullableText(r.citation_ou_element);
      return {
        niveau: NIVEAUX.indexOf(r.niveau) !== -1 ? r.niveau : 'Recommandé',
        titre: nullableText(r.titre) || 'Remarque sans titre',
        pages,
        element,
        // true : passage retrouvé mot pour mot ; sinon il peut s'agir d'un élément visuel décrit
        citationRetrouvee: element ? checkCitation(element, pages, pageTexts) === true : false,
        explication: nullableText(r.explication) || '',
        proposition: nullableText(r.proposition),
        reference: nullableText(r.reference),
        confiance: CONFIANCES.indexOf(r.confiance) !== -1 ? r.confiance : 'moyen',
        ordre: i
      };
    });
    // "À traiter" d'abord ; l'ordre d'importance donné par l'IA est conservé dans chaque niveau
    remarques.sort((x, y) => NIVEAUX.indexOf(x.niveau) - NIVEAUX.indexOf(y.niveau) || x.ordre - y.ordre);
    remarques.forEach(r => { delete r.ordre; });
    const valeur = APPRECIATIONS.indexOf(app.valeur) !== -1 ? app.valeur : null;
    const counts = {
      aTraiter: remarques.filter(r => r.niveau === 'À traiter').length,
      recommandes: remarques.filter(r => r.niveau === 'Recommandé').length,
      suggestions: remarques.filter(r => r.niveau === 'Suggestion').length
    };
    return {
      comprehension: {
        nature: nullableText(comp.nature) || '',
        metEnAvant: nullableText(comp.met_en_avant) || '',
        canal: nullableText(comp.canal_probable) || '',
        objectif: nullableText(comp.objectif) || '',
        resume: nullableText(comp.resume) || '',
        ecart: nullableText(comp.ecart_declaration)
      },
      appreciation: { valeur, synthese: nullableText(app.synthese) || '' },
      score: scoreFromAppreciation(valeur),
      pointsForts: textList(a.points_forts),
      remarques,
      pointsAVerifier: textList(a.points_a_verifier),
      counts
    };
  }

  // Relevé de l'étape 1 : une entrée par page du document, dans l'ordre, pages manquantes signalées
  function mergeStep1(results, totalPages) {
    const byPage = {};
    results.forEach(r => (Array.isArray(r && r.pages) ? r.pages : []).forEach(p => {
      if (p && Number.isInteger(p.numero) && p.numero >= 1 && p.numero <= totalPages && !byPage[p.numero]) byPage[p.numero] = p;
    }));
    const releve = [];
    const manquantes = [];
    for (let i = 1; i <= totalPages; i++) {
      if (byPage[i]) releve.push(byPage[i]);
      else manquantes.push(i);
    }
    return { releve, manquantes };
  }

  // ---------------------------------------------------------------------------
  // Rapport PDF (jsPDF, police standard) : caractères hors Windows-1252 remplacés
  // ---------------------------------------------------------------------------
  const WINANSI_EXTRA = '€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ';
  function pdfSafe(s) {
    return toText(s).normalize('NFC').replace(/[‐‑‒−]/g, '-').replace(/ | /g, ' ')
      .split(/(?:)/u).map(ch => {
        const code = ch.codePointAt(0);
        if (code === 10 || (code >= 32 && code <= 126) || (code >= 160 && code <= 255) || WINANSI_EXTRA.indexOf(ch) !== -1) return ch;
        return '?';
      }).join('');
  }

  // Couleurs partagées par l'affichage et le rapport PDF ("À traiter" en orange)
  const COULEUR_NIVEAU = { 'À traiter': [194, 65, 12], 'Recommandé': [29, 78, 216], 'Suggestion': [71, 85, 105] };
  const COULEUR_APPRECIATION = { 'Prêt à soumettre': [30, 126, 74], 'Quelques ajustements conseillés': [183, 119, 13], 'À retravailler': [194, 65, 12] };

  function generateReportPdf(jsPDF, report) {
    const doc = new jsPDF({ unit: 'mm', format: 'a4' });
    const W = 210, M = 16, maxW = W - 2 * M;
    let y = 0;
    const newPageIfNeeded = h => { if (y + h > 282) { doc.addPage(); y = 18; } };
    const text = (s, size, style, color, indent) => {
      doc.setFont('helvetica', style || 'normal'); doc.setFontSize(size); doc.setTextColor.apply(doc, color || [40, 40, 40]);
      const lines = doc.splitTextToSize(pdfSafe(s), maxW - (indent || 0));
      lines.forEach(line => { newPageIfNeeded(size * 0.45); doc.text(line, M + (indent || 0), y); y += size * 0.45; });
    };

    doc.setFillColor(0, 20, 59); doc.rect(0, 0, W, 30, 'F');
    doc.setTextColor(209, 143, 65); doc.setFont('helvetica', 'bold'); doc.setFontSize(15); doc.text('EIFFEL INVESTMENT GROUP', M, 13);
    doc.setTextColor(255, 255, 255); doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.text(pdfSafe('Doc Review : relecture préalable de la communication commerciale'), M, 21);
    y = 40;

    const comp = report.comprehension || {};
    const app = report.appreciation || {};
    const meta = [
      ['Référence', report.ref], ['Document', report.nomFichier], ['Type de fonds', report.typeFonds], ['Public visé', report.marcheLibelle],
      ['Nature détectée', comp.nature || '-'], ['Appréciation', app.valeur || '-'],
      ['Empreinte SHA-256', report.hashFichier || '-'], ['Date', report.dateLibelle], ['Utilisateur', report.demandeur],
      ['Modèles', report.modeles], ['Pages analysées', String(report.totalPages)]
    ];
    meta.forEach(([k, v]) => {
      newPageIfNeeded(6);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); doc.setTextColor(0, 20, 59); doc.text(pdfSafe(k + ' :'), M, y);
      doc.setFont('helvetica', 'normal'); doc.setTextColor(50, 50, 50);
      const lines = doc.splitTextToSize(pdfSafe(v), maxW - 42);
      lines.forEach((l, i) => { if (i) { y += 4.2; newPageIfNeeded(5); } doc.text(l, M + 42, y); });
      y += 5.2;
    });

    const c = report.counts || {};
    const section = title => { y += 3; newPageIfNeeded(10); text(title, 10, 'bold', [0, 20, 59]); };
    const bullets = list => { if (!list.length) text('-', 9, 'normal', null, 2); list.forEach(s => text('- ' + s, 9, 'normal', null, 2)); };

    y += 3;
    text(`Appréciation globale : ${app.valeur || '-'}`, 11.5, 'bold', COULEUR_APPRECIATION[app.valeur] || [0, 20, 59]);
    text(`À traiter : ${c.aTraiter || 0}   -   Recommandé : ${c.recommandes || 0}   -   Suggestion : ${c.suggestions || 0}`, 9.5, 'normal', [70, 70, 70]);
    if ((report.avertissements || []).length) {
      y += 2;
      text('À noter', 10, 'bold', [146, 64, 14]);
      report.avertissements.forEach(a => text('- ' + a, 9, 'normal', [146, 64, 14], 2));
    }
    section('Compréhension du document');
    text(comp.resume || '-', 9, 'normal');
    if (comp.metEnAvant) text('Met en avant : ' + comp.metEnAvant, 9, 'normal', [80, 80, 80], 2);
    if (comp.canal) text('Canal probable : ' + comp.canal, 9, 'normal', [80, 80, 80], 2);
    if (comp.objectif) text('Objectif : ' + comp.objectif, 9, 'normal', [80, 80, 80], 2);
    if (comp.ecart) text('Écart avec les informations déclarées : ' + comp.ecart, 9, 'bold', [146, 64, 14], 2);
    section('Synthèse');
    text(app.synthese || '-', 9, 'normal');
    section('Points forts');
    bullets(report.pointsForts || []);
    section('Remarques');
    if (!(report.remarques || []).length) text('Aucune remarque.', 9, 'normal');
    (report.remarques || []).forEach((r, i) => {
      y += 2;
      newPageIfNeeded(12);
      text(`${i + 1}. [${r.niveau.toUpperCase()}] ${r.titre}` + (r.pages.length ? `  (pages ${r.pages.join(', ')})` : '  (document entier)'), 9.5, 'bold', COULEUR_NIVEAU[r.niveau]);
      if (r.element) text((r.citationRetrouvee ? 'Passage cité : "' + r.element + '"' : 'Passage ou élément : ' + r.element), 9, 'italic', [80, 80, 80], 3);
      if (r.explication) text(r.explication, 9, 'normal', [50, 50, 50], 3);
      if (r.proposition) text('Proposition : ' + r.proposition, 9, 'normal', [0, 102, 96], 3);
      if (r.reference) text('Référence : ' + r.reference, 9, 'normal', [80, 80, 80], 3);
      text('Confiance : ' + r.confiance, 8.5, 'normal', [120, 120, 120], 3);
    });
    section('Points à vérifier avant soumission');
    bullets(report.pointsAVerifier || []);

    const pages = doc.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(140, 140, 140);
      doc.text(pdfSafe(`${report.ref} - aide à la relecture, ne vaut pas validation Compliance - page ${p}/${pages}`), W / 2, 292, { align: 'center' });
    }
    return doc.output('arraybuffer');
  }

  async function sha256Hex(buffer) {
    const digest = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, '0')).join('');
  }

  return Object.freeze({
    ANALYSE_CONFIG, APPRECIATIONS, SCORE_PAR_APPRECIATION, NIVEAUX, COULEUR_NIVEAU, COULEUR_APPRECIATION,
    escapeHtml, scoreFromAppreciation, isProfessionalFundForRetail, makeRef, sanitizeFileName, extensionOf, baseName,
    normalizeForMatch, checkCitation, buildBatches, parsePptxStructure, mapPdfPages,
    consolidate, mergeStep1, pdfSafe, generateReportPdf, sha256Hex
  });
})();
