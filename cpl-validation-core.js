// =============================================
//  VALIDATION DOCUMENTAIRE : marquage PDF (logique pure)
//  Compliance App — Eiffel Investment Group
// =============================================
// Aucune dépendance au DOM ni à Microsoft Graph : pdf-lib et pdf.js sont passés en
// paramètre. Marquage de chaque page ("Réf. CPL-AAAA-NNNN · Validé le JJ/MM/AAAA"),
// en bas à droite de la zone visible (CropBox), en tenant compte de la rotation de
// la page, puis contrôle par relecture.

const CplValidationCore = (() => {
  'use strict';

  const MARK = Object.freeze({ taillePt: 7, margePt: 14, gris: 0.45 });

  function pad2(n) { return String(n).padStart(2, '0'); }

  function buildLabel(ref, date) {
    return `Réf. ${ref} · Validé le ${pad2(date.getDate())}/${pad2(date.getMonth() + 1)}/${date.getFullYear()}`;
  }

  function sanitizeFileName(name) {
    let n = String(name || '').normalize('NFC').replace(/["*:<>?/\\|\u0000-\u001f]/g, '_').replace(/\s+/g, ' ').trim();
    n = n.replace(/^[.\s~$]+/, '').replace(/[.\s]+$/, '');
    return n || 'document';
  }

  function baseName(name) {
    const n = String(name || '');
    const dot = n.lastIndexOf('.');
    return dot > 0 ? n.slice(0, dot) : n;
  }

  function extensionOf(name) {
    const m = /\.([^.]+)$/.exec(String(name || ''));
    return m ? m[1].toLowerCase() : '';
  }

  // "<nom d'origine> - VALIDE CPL-2026-0123.pdf"
  function markedFileName(originalName, ref) {
    return sanitizeFileName(`${baseName(originalName)} - VALIDE ${ref}.pdf`);
  }

  function isMarkedFileName(name) { return / - VALIDE CPL-\d{4}-\d{4,}( \d+)?\.(pdf|pptx|docx)$/i.test(String(name || '')); }

  // "<nom d'origine> - VALIDE CPL-2026-0123.pptx" (fichier d'origine marqué, extension conservée)
  function markedOriginalName(originalName, ref) {
    const ext = extensionOf(originalName);
    return sanitizeFileName(`${baseName(originalName)} - VALIDE ${ref}.${ext}`);
  }

  // ---------------------------------------------------------------------------
  // Marquage du fichier d'origine (option) : PPTX et DOCX modifiés par JSZip.
  // Le fichier marqué est enregistré sous un nouveau nom et contrôlé par reconversion
  // en PDF ; l'original n'est jamais modifié. Le PDF marqué reste la version officielle.
  // ---------------------------------------------------------------------------
  const NS = {
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  };
  const REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';
  const REL_FOOTER = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer';

  function xmlEscape(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

  function parseXml(text, label) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error(`XML illisible (${label})`);
    return doc;
  }

  function serializeXml(doc, originalText) {
    const body = new XMLSerializer().serializeToString(doc);
    const decl = /^\s*<\?xml[^>]*\?>/.exec(originalText);
    return (decl ? decl[0] + '\r\n' : '') + body.replace(/^<\?xml[^>]*\?>\s*/, '');
  }

  function resolvePart(baseDir, target) {
    const parts = (target.startsWith('/') ? target.slice(1) : baseDir + '/' + target).split('/');
    const out = [];
    parts.forEach(p => { if (p === '..') out.pop(); else if (p && p !== '.') out.push(p); });
    return out.join('/');
  }

  function relationships(relsDoc, type) {
    const map = {};
    Array.from(relsDoc.getElementsByTagNameNS(NS.rel, 'Relationship')).forEach(r => {
      if (!type || r.getAttribute('Type') === type) map[r.getAttribute('Id')] = r.getAttribute('Target');
    });
    return map;
  }

  async function markPptx(JSZip, bytes, label) {
    const zip = await JSZip.loadAsync(bytes);
    const presFile = zip.file('ppt/presentation.xml');
    const relsFile = zip.file('ppt/_rels/presentation.xml.rels');
    if (!presFile || !relsFile) throw new Error('structure PowerPoint invalide');
    const pres = parseXml(await presFile.async('string'), 'presentation.xml');
    const slideRels = relationships(parseXml(await relsFile.async('string'), 'presentation.xml.rels'), REL_SLIDE);
    const size = pres.getElementsByTagNameNS(NS.p, 'sldSz')[0];
    const W = size ? Number(size.getAttribute('cx')) : 12192000;
    const H = size ? Number(size.getAttribute('cy')) : 6858000;
    const cx = Math.min(3400000, Math.round(W * 0.6));
    const cy = 190000;
    const x = Math.max(0, W - cx - 150000);
    const y = Math.max(0, H - cy - 90000);
    const ids = Array.from(pres.getElementsByTagNameNS(NS.p, 'sldId'));
    if (!ids.length) throw new Error('présentation sans slide');
    for (const sldId of ids) {
      const path = resolvePart('ppt', slideRels[sldId.getAttributeNS(NS.r, 'id')] || '');
      const f = zip.file(path);
      if (!f) throw new Error('slide introuvable : ' + path);
      const text = await f.async('string');
      const doc = parseXml(text, path);
      const tree = doc.getElementsByTagNameNS(NS.p, 'spTree')[0];
      if (!tree) throw new Error('slide sans arbre de formes : ' + path);
      let maxId = 0;
      Array.from(doc.getElementsByTagNameNS('*', 'cNvPr')).forEach(n => { const v = parseInt(n.getAttribute('id'), 10); if (v > maxId) maxId = v; });
      const shapeXml = `<p:sp xmlns:p="${NS.p}" xmlns:a="${NS.a}"><p:nvSpPr><p:cNvPr id="${maxId + 1}" name="Validation CPL"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
        + `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>`
        + `<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="b"><a:noAutofit/></a:bodyPr><a:lstStyle/>`
        + `<a:p><a:pPr algn="r"/><a:r><a:rPr lang="fr-FR" sz="700" dirty="0"><a:solidFill><a:srgbClr val="737373"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${xmlEscape(label)}</a:t></a:r></a:p></p:txBody></p:sp>`;
      const shape = doc.importNode(parseXml(shapeXml, 'forme').documentElement, true);
      const extLst = Array.from(tree.childNodes).find(n => n.nodeType === 1 && n.localName === 'extLst');
      tree.insertBefore(shape, extLst || null);
      zip.file(path, serializeXml(doc, text));
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  async function markDocx(JSZip, bytes, label) {
    const zip = await JSZip.loadAsync(bytes);
    const docFile = zip.file('word/document.xml');
    const relsFile = zip.file('word/_rels/document.xml.rels');
    if (!docFile || !relsFile) throw new Error('structure Word invalide');
    const doc = parseXml(await docFile.async('string'), 'document.xml');
    const footerRels = relationships(parseXml(await relsFile.async('string'), 'document.xml.rels'), REL_FOOTER);
    const refs = Array.from(doc.getElementsByTagNameNS(NS.w, 'footerReference')).map(n => n.getAttributeNS(NS.r, 'id'));
    const parts = [];
    refs.forEach(rid => { const t = footerRels[rid]; if (t) { const p = resolvePart('word', t); if (parts.indexOf(p) === -1) parts.push(p); } });
    if (!parts.length) throw new Error('le document Word n\'a pas de pied de page');
    for (const path of parts) {
      const f = zip.file(path);
      if (!f) throw new Error('pied de page introuvable : ' + path);
      const text = await f.async('string');
      const ftr = parseXml(text, path);
      const root = ftr.documentElement;
      if (root.localName !== 'ftr') throw new Error('pied de page inattendu : ' + path);
      const pXml = `<w:p xmlns:w="${NS.w}"><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:color w:val="737373"/><w:sz w:val="14"/><w:szCs w:val="14"/></w:rPr><w:t xml:space="preserve">${xmlEscape(label)}</w:t></w:r></w:p>`;
      root.appendChild(ftr.importNode(parseXml(pXml, 'paragraphe').documentElement, true));
      zip.file(path, serializeXml(ftr, text));
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  // Détection d'une signature électronique (le marquage la rend invalide)
  function looksSigned(bytes) {
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    let s = '';
    for (let i = 0; i < b.length; i += 65536) s += String.fromCharCode.apply(null, b.subarray(i, Math.min(i + 65536, b.length)));
    return /\/ByteRange\s*\[/.test(s) && /\/Type\s*\/Sig\b/.test(s);
  }

  // Position et angle du texte pour qu'il apparaisse en bas à droite, dans le sens de lecture,
  // quelle que soit la rotation de la page (0, 90, 180, 270 degrés).
  function placement(crop, rotation, textWidth, margin) {
    const x0 = crop.x, y0 = crop.y, x1 = crop.x + crop.width, y1 = crop.y + crop.height;
    switch (rotation) {
      case 90: return { x: x1 - margin, y: y1 - margin - textWidth, angle: 90 };
      case 180: return { x: x0 + margin + textWidth, y: y1 - margin, angle: 180 };
      case 270: return { x: x0 + margin, y: y0 + margin + textWidth, angle: 270 };
      default: return { x: x1 - margin - textWidth, y: y0 + margin, angle: 0 };
    }
  }

  function normalizeRotation(angle) {
    const a = ((Math.round((Number(angle) || 0) / 90) * 90) % 360 + 360) % 360;
    return a;
  }

  async function markPdf(PDFLib, bytes, label) {
    const { PDFDocument, StandardFonts, rgb, degrees, pushGraphicsState, popGraphicsState } = PDFLib;
    let doc;
    try {
      doc = await PDFDocument.load(bytes, { updateMetadata: false });
    } catch (e) {
      if (/encrypt/i.test(String(e && (e.name + ' ' + e.message)))) throw new Error('PDF protégé ou chiffré : marquage impossible.');
      throw new Error('PDF illisible : ' + (e && e.message));
    }
    if (doc.isEncrypted) throw new Error('PDF protégé ou chiffré : marquage impossible.');
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const size = MARK.taillePt;
    const textWidth = font.widthOfTextAtSize(label, size);
    const pages = doc.getPages();
    if (!pages.length) throw new Error('PDF sans page.');
    const ctx = doc.context;
    pages.forEach(page => {
      // Le contenu existant est isolé (q ... Q) : la mention ne dépend pas de l'état graphique
      // laissé par le document (transformation, couleur, découpe).
      const start = ctx.register(ctx.contentStream([pushGraphicsState()]));
      const end = ctx.register(ctx.contentStream([popGraphicsState()]));
      page.node.wrapContentStreams(start, end);
      const crop = page.getCropBox();
      const rotation = normalizeRotation(page.getRotation().angle);
      const p = placement(crop, rotation, textWidth, MARK.margePt);
      page.drawText(label, { x: p.x, y: p.y, size, font, color: rgb(MARK.gris, MARK.gris, MARK.gris), rotate: degrees(p.angle) });
    });
    const out = await doc.save();
    return { bytes: out, pageCount: pages.length };
  }

  function compact(s) { return String(s || '').normalize('NFC').replace(/\s+/g, ''); }

  async function pageTexts(pdfjsLib, bytes) {
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(bytes).slice(), isEvalSupported: false }).promise;
    const texts = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const tc = await (await pdf.getPage(i)).getTextContent();
      texts.push(tc.items.map(it => it.str || '').join(''));
    }
    return texts;
  }

  // Contrôle : même nombre de pages que le PDF source, mention présente sur chaque page
  async function verifyMarkedPdf(pdfjsLib, bytes, label, expectedPages) {
    const texts = await pageTexts(pdfjsLib, bytes);
    if (texts.length !== expectedPages) throw new Error(`le PDF marqué compte ${texts.length} page(s) au lieu de ${expectedPages}`);
    const target = compact(label);
    const missing = [];
    texts.forEach((t, i) => { if (compact(t).indexOf(target) === -1) missing.push(i + 1); });
    if (missing.length) throw new Error('mention absente des pages ' + missing.join(', '));
    return { pages: texts.length };
  }

  function containsLabel(texts, ref) {
    const target = compact('Réf. ' + ref);
    return texts.some(t => compact(t).indexOf(target) !== -1);
  }

  return Object.freeze({
    MARK, buildLabel, markedFileName, markedOriginalName, isMarkedFileName, extensionOf, looksSigned,
    placement, normalizeRotation, markPdf, pageTexts, verifyMarkedPdf, containsLabel, markPptx, markDocx
  });
})();
