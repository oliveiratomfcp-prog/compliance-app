// =============================================
//  DOC REVIEW : options du formulaire
//  Compliance App — Eiffel Investment Group
// =============================================
// Chargé par doc-review.html. Pour ajouter, retirer ou renommer un type de fonds,
// modifier la liste ci-dessous (un libellé par ligne). Le libellé choisi est transmis
// tel quel à l'IA et enregistré dans la colonne TypeFonds de "Doc Review Analyses".
// Contraintes : 120 caractères au plus, sans accolades ni retour à la ligne.

const DOC_REVIEW_TYPES_FONDS = Object.freeze([
  'Aucun fonds en particulier (communication institutionnelle)',
  'OPCVM (UCITS)',
  'FIVG (fonds d\'investissement à vocation générale)',
  'FCPR',
  'FCPI',
  'FIP',
  'FPCI (fonds professionnel de capital investissement)',
  'FPS (fonds professionnel spécialisé)',
  'SLP (société de libre partenariat)',
  'ELTIF',
  'Organisme de titrisation (FCT, OT)',
  'Fonds de droit étranger (ex. SICAV-RAIF, SCSp)',
  'Plusieurs fonds ou gamme de fonds',
  'Autre'
]);

// Types de fonds en principe réservés aux investisseurs professionnels : un avertissement
// non bloquant s'affiche si le public Retail est choisi. Libellés identiques à la liste ci-dessus.
const DOC_REVIEW_FONDS_PROFESSIONNELS = Object.freeze([
  'FPCI (fonds professionnel de capital investissement)',
  'FPS (fonds professionnel spécialisé)',
  'SLP (société de libre partenariat)'
]);
