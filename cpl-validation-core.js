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

  function isMarkedFileName(name) { return / - VALIDE CPL-\d{4}-\d{4,}( \d+)?\.pdf$/i.test(String(name || '')); }

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
    MARK, buildLabel, markedFileName, isMarkedFileName, extensionOf, looksSigned,
    placement, normalizeRotation, markPdf, pageTexts, verifyMarkedPdf, containsLabel
  });
})();
