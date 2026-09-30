// =============================================
//  PROFIL COMPLIANCE — configuration partagée
//  Compliance App — Eiffel Investment Group
// =============================================
// Chargé par index.html, tracker-cpl.html et counterflow.html.
// Ce contrôle ne sert qu'à l'affichage : la vraie protection repose sur les
// permissions SharePoint des listes concernées.

const CPL_EMAILS = Object.freeze([
  'thomas.oliveira@eiffel-ig.com',
  'cassandre.memet@eiffel-ig.com',
  'manalle.ourrami@eiffel-ig.com',
  'jessica.manukyan@eiffel-ig.com'
]);

// true si l'email appartient à l'équipe Compliance (comparaison insensible à la casse)
function isCPLUser(email) {
  if (typeof email !== 'string') return false;
  const e = email.toLowerCase();
  return CPL_EMAILS.some(c => c.toLowerCase() === e);
}
