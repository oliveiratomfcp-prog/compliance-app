// =============================================
//  DOC REVIEW : logique pure (navigateur)
//  Compliance App — Eiffel Investment Group
// =============================================
// Aucune dépendance au DOM, à MSAL ni à Microsoft Graph : score, contrôle des
// citations, structure PPTX, correspondance pages / slides, regroupement des
// constats, rapport PDF. Testable seul.

const DocReviewCore = (() => {
  'use strict';

  // Pénalités du score (calculé par le code, jamais par l'IA) : 100 moins une pénalité
  // par constat bloquant et par point d'attention, minimum 0.
  const SCORE_CONFIG = Object.freeze({ penaliteBloquant: 20, penaliteAttention: 5 });

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

  // Doit rester identique à la liste serveur (api/shared/docreview-logic.js)
  const TYPES_DOCUMENT = Object.freeze([
    'présentation commerciale', 'fiche produit / DICI', 'email marketing', 'post LinkedIn',
    'post réseaux sociaux', 'rapport de gestion', 'document publicitaire', 'communiqué de presse'
  ]);
  const GRAVITES = ['bloquant', 'attention', 'conforme'];

  function toText(v) { return v === null || v === undefined ? '' : String(v); }

  function escapeHtml(v) {
    return toText(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function computeScore(counts, config) {
    const c = config || SCORE_CONFIG;
    return Math.max(0, 100 - c.penaliteBloquant * (counts.bloquants || 0) - c.penaliteAttention * (counts.attention || 0));
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
  // Consolidation : constats finaux, contrôle des pages et des citations, score
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

  function consolidate(synthese, pageTexts, totalPages, config) {
    const constats = (Array.isArray(synthese && synthese.constats) ? synthese.constats : []).map(c => {
      const pages = cleanPages(c.pages, totalPages);
      const citation = nullableText(c.citation);
      return {
        gravite: GRAVITES.indexOf(c.gravite) !== -1 ? c.gravite : 'attention',
        nature: toText(c.nature) || 'non_conforme',
        titre: nullableText(c.titre) || 'Constat sans titre',
        explication: nullableText(c.explication) || '',
        reference: nullableText(c.reference),
        citation,
        citationVerifiee: citation ? checkCitation(citation, pages, pageTexts) : null,
        description_visuelle: nullableText(c.description_visuelle),
        correction: nullableText(c.correction),
        incertitude: nullableText(c.incertitude),
        pages
      };
    });
    const order = { bloquant: 0, attention: 1, conforme: 2 };
    constats.sort((a, b) => order[a.gravite] - order[b.gravite] || (a.pages[0] || 0) - (b.pages[0] || 0));
    const counts = {
      bloquants: constats.filter(c => c.gravite === 'bloquant').length,
      attention: constats.filter(c => c.gravite === 'attention').length,
      conformes: constats.filter(c => c.gravite === 'conforme').length
    };
    return {
      constats,
      counts,
      score: computeScore(counts, config),
      synthese: nullableText(synthese && synthese.synthese) || '',
      formatDetecte: nullableText(synthese && synthese.format_detecte) || '',
      recommandation: toText(synthese && synthese.recommandation),
      pagesSansAvertissement: cleanPages(synthese && synthese.pages_sans_avertissement, totalPages),
      themes: (Array.isArray(synthese && synthese.themes) ? synthese.themes : []).map(toText).filter(Boolean)
    };
  }

  // Regroupement pour l'affichage : constats du document entier, puis par première page citée
  function groupByPage(constats) {
    const groups = [];
    const general = constats.filter(c => c.pages.length === 0);
    if (general.length) groups.push({ page: null, constats: general });
    const byPage = {};
    constats.filter(c => c.pages.length > 0).forEach(c => { (byPage[c.pages[0]] = byPage[c.pages[0]] || []).push(c); });
    Object.keys(byPage).map(Number).sort((a, b) => a - b).forEach(p => groups.push({ page: p, constats: byPage[p] }));
    return groups;
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

  const GRAVITE_LIBELLE = { bloquant: 'BLOQUANT', attention: 'ATTENTION', conforme: 'CONFORME' };

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
    doc.setTextColor(255, 255, 255); doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.text(pdfSafe('Doc Review : rapport de pré-contrôle AMF / MIF II'), M, 21);
    y = 40;

    const meta = [
      ['Référence', report.ref], ['Document', report.nomFichier], ['Type', report.typeDocument], ['Marché', report.marcheLibelle],
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

    y += 3;
    text(`Score : ${report.score} / 100   -   Bloquants : ${report.counts.bloquants}   -   Points d'attention : ${report.counts.attention}   -   Conformes : ${report.counts.conformes}`, 11, 'bold', [0, 20, 59]);
    y += 2;
    if (report.avertissements.length) {
      text('Avertissements', 10, 'bold', [146, 64, 14]);
      report.avertissements.forEach(a => text('- ' + a, 9, 'normal', [146, 64, 14], 2));
      y += 2;
    }
    text('Synthèse', 10, 'bold', [0, 20, 59]);
    text(report.synthese || '-', 9, 'normal');
    y += 3;
    text('Constats', 10, 'bold', [0, 20, 59]);
    report.constats.forEach((c, i) => {
      y += 2;
      newPageIfNeeded(12);
      const color = c.gravite === 'bloquant' ? [192, 57, 43] : c.gravite === 'attention' ? [183, 119, 13] : [30, 126, 74];
      text(`${i + 1}. [${GRAVITE_LIBELLE[c.gravite]}] ${c.titre}` + (c.pages.length ? `  (pages ${c.pages.join(', ')})` : '  (document entier)'), 9.5, 'bold', color);
      if (c.explication) text(c.explication, 9, 'normal', [50, 50, 50], 3);
      if (c.citation) text('Citation : "' + c.citation + '"' + (c.citationVerifiee === false ? '  [citation non retrouvée dans le texte extrait]' : ''), 9, 'italic', [80, 80, 80], 3);
      if (c.description_visuelle) text('Élément visuel : ' + c.description_visuelle, 9, 'normal', [80, 80, 80], 3);
      if (c.reference) text('Référence : ' + c.reference, 9, 'normal', [80, 80, 80], 3);
      if (c.incertitude) text('Incertitude : ' + c.incertitude, 9, 'normal', [146, 64, 14], 3);
      if (c.correction) text('Correction suggérée : ' + c.correction, 9, 'normal', [0, 102, 96], 3);
    });

    const pages = doc.getNumberOfPages();
    for (let p = 1; p <= pages; p++) {
      doc.setPage(p);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(140, 140, 140);
      doc.text(pdfSafe(`${report.ref} - pré-contrôle automatisé, ne vaut pas validation Compliance - page ${p}/${pages}`), W / 2, 292, { align: 'center' });
    }
    return doc.output('arraybuffer');
  }

  async function sha256Hex(buffer) {
    const digest = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, '0')).join('');
  }

  return Object.freeze({
    SCORE_CONFIG, ANALYSE_CONFIG, TYPES_DOCUMENT,
    escapeHtml, computeScore, makeRef, sanitizeFileName, extensionOf, baseName,
    normalizeForMatch, checkCitation, buildBatches, parsePptxStructure, mapPdfPages,
    consolidate, groupByPage, mergeStep1, pdfSafe, generateReportPdf, sha256Hex
  });
})();
