// =============================================
//  TRACKER CPL : traduction de l'interface FR / EN
//  Compliance App — Eiffel Investment Group
// =============================================
// Chargé par tracker-cpl.html. L'interface est écrite en français ; en anglais, les textes
// affichés sont traduits à la volée (MutationObserver) à partir du dictionnaire et des
// motifs ci-dessous, puis restaurés à l'identique au retour en français.
// Ne sont jamais traduits : les zones éditables, les éléments marqués .no-i18n (description,
// commentaires), les valeurs envoyées à SharePoint et les emails de notification.
// Pour traduire un nouveau texte : ajouter la ligne "texte français": "English text" dans
// TRACKER_EN (texte exact, espaces multiples ramenés à un seul) ou un motif dans
// TRACKER_EN_PATTERNS s'il contient une partie variable.

const TRACKER_EN = {
  // Connexion, barre latérale
  "Connectez-vous avec votre compte d'entreprise Microsoft pour accéder au portail de demandes Compliance.": "Sign in with your Microsoft corporate account to access the Compliance request portal.",
  "Se connecter avec Microsoft": "Sign in with Microsoft",
  "Accès réservé aux collaborateurs Eiffel IG": "Access restricted to Eiffel IG employees",
  "Déclarer": "Declare",
  "Chargement...": "Loading...",
  "Collaborateur": "Employee",
  "Se déconnecter": "Sign out",
  "Chargement depuis SharePoint...": "Loading from SharePoint...",

  // Liste des demandes
  "Suivi des demandes · Compliance": "Request tracking · Compliance",
  "Mes demandes": "My requests",
  "Toutes les demandes": "All requests",
  "Nouvelle demande": "New request",
  "Demandes": "Requests",
  "Dashboard KPI": "KPI Dashboard",
  "Exporter CSV": "Export CSV",
  "Colonnes": "Columns",
  "Colonnes visibles": "Visible columns",
  "Réinitialiser": "Reset",
  "✕ Réinitialiser": "✕ Reset",
  "Mes demandes assignées": "My assigned requests",
  "Tous": "All",
  "(vide)": "(empty)",
  "↩ Action requise": "↩ Action required",
  "Aucune demande dans cette catégorie": "No requests in this category",
  "Fermer": "Close",
  "Filtrer…": "Filter…",
  "Personnaliser les colonnes": "Customize columns",
  "Rechercher...": "Search...",
  "Rafraîchir la liste depuis SharePoint": "Refresh the list from SharePoint",
  "Liste actualisée.": "List refreshed.",

  // Colonnes et champs
  "Référence": "Reference",
  "Objet": "Subject",
  "Type": "Type",
  "Demandeur": "Requester",
  "Département": "Department",
  "Priorité": "Priority",
  "Statut": "Status",
  "Soumis le": "Submitted on",
  "Soumis": "Submitted",
  "Deadline": "Deadline",
  "Responsable CPL": "CPL owner",
  "Description": "Description",
  "Activité": "Activity",

  // Statuts
  "À traiter": "To do",
  "Réponse à revoir": "Reply to review",
  "En cours d'analyse": "Under review",
  "En attente d'informations": "Awaiting information",
  "Traité": "Processed",
  "Clôturé": "Closed",
  "Rejeté": "Rejected",
  "Résolu": "Resolved",

  // Priorités
  "Faible": "Low",
  "Élevé": "High",
  "Critique": "Critical",

  // Types de demande et départements
  "Validation documentation commerciale": "Marketing documentation approval",
  "Appel d'offre / DDQ": "Tender / DDQ",
  "Avis réglementaire": "Regulatory opinion",
  "Procédures": "Procedures",
  "Conflit d'intérêts": "Conflict of interest",
  "Déontologie": "Ethics",
  "Processus d'investissement": "Investment process",
  "Franchissement de seuil": "Threshold crossing",
  "Juridique": "Legal",

  // Nouvelle demande
  "Département *": "Department *",
  "Sélectionner un département...": "Select a department...",
  "Type de demande *": "Request type *",
  "Sélectionner un type...": "Select a type...",
  "Priorité *": "Priority *",
  "⚠ Une deadline est obligatoire pour cette priorité.": "⚠ A deadline is required for this priority.",
  "Objet *": "Subject *",
  "— 150 car. max": "(150 characters max)",
  "Personnes en copie": "People in copy",
  "(optionnel)": "(optional)",
  "Description *": "Description *",
  "Deadline souhaitée": "Requested deadline",
  "Deadline *": "Deadline *",
  "(obligatoire)": "(required)",
  "Pièces jointes": "Attachments",
  "Cliquez ou déposez vos fichiers ici": "Click or drop your files here",
  "PDF, Word, Excel, images — plusieurs fichiers acceptés": "PDF, Word, Excel, images (multiple files accepted)",
  "Soumettre la demande": "Submit request",
  "Envoi en cours...": "Sending...",
  "Copie des fichiers Doc Review...": "Copying Doc Review files...",
  "Titre court et descriptif...": "Short descriptive title...",
  "Tapez un nom pour rechercher...": "Type a name to search...",
  "Gras": "Bold",
  "Italique": "Italic",
  "Souligné": "Underline",
  "Liste": "List",
  "Décrivez votre demande en détail (contexte, enjeux, documents disponibles)...": "Describe your request in detail (context, stakes, available documents)...",
  "Retirer": "Remove",
  "Ajouter une personne en copie...": "Add a person in copy...",
  "Veuillez remplir tous les champs obligatoires (*)": "Please fill in all required fields (*)",
  "Une deadline est obligatoire pour les priorités Élevé et Critique.": "A deadline is required for High and Critical priorities.",

  // Détail d'une demande
  "Changer le statut": "Change status",
  "Sélectionner un statut...": "Select a status...",
  "Motif de rejet *": "Rejection reason *",
  "Expliquez la raison du rejet...": "Explain the reason for rejection...",
  "Appliquer": "Apply",
  "Attribuer à": "Assign to",
  "Sélectionner...": "Select...",
  "Attribuer": "Assign",
  "Actuellement": "Currently",
  "Validation documentaire": "Document validation",
  "Valider et marquer": "Validate and stamp",
  "Génère un PDF portant la référence de la demande sur chaque page et passe la demande au statut Traité.": "Generates a PDF bearing the request reference on each page and moves the request to Processed status.",
  "Rouvrir cette demande": "Reopen this request",
  "Rouvrir cette demande ?": "Reopen this request?",
  "Enregistré.": "Saved.",
  "Enregistrement...": "Saving...",
  "Mettre à jour": "Update",
  "Déposer": "Upload",
  "Upload...": "Uploading...",
  "Commentaires": "Comments",
  "Joindre": "Attach",
  "Note interne (Compliance uniquement)": "Internal note (Compliance only)",
  "Note interne": "Internal note",
  "Envoyer": "Send",
  "Envoi...": "Sending...",
  "Ajouter un commentaire...": "Add a comment...",
  "Historique des statuts": "Status history",
  "Aucun fichier déposé.": "No files uploaded.",
  "Télécharger": "Download",
  "Ouvrir pour commenter": "Open to comment",
  "Ouvrir": "Open",
  "Dossier introuvable sur SharePoint.": "Folder not found on SharePoint.",
  "Liste commentaires non configurée.": "Comments list not configured.",
  "Aucun commentaire.": "No comments.",
  "Erreur de chargement des commentaires.": "Error loading comments.",
  "Document validé et marqué": "Document validated and stamped",
  "Finaliser la liaison Doc Review (joindre les fichiers)": "Finalize Doc Review link (attach files)",
  "Veuillez saisir un commentaire.": "Please enter a comment.",
  "Message envoyé, mais le statut « Réponse à revoir » n'a pas pu être appliqué.": "Message sent, but the \"Reply to review\" status could not be applied.",
  "Veuillez sélectionner un statut.": "Please select a status.",
  "Le motif de rejet est obligatoire.": "A rejection reason is required.",
  "Demande rouverte, statut repassé à Soumis.": "Request reopened, status set back to Submitted.",
  "Veuillez sélectionner un membre CPL.": "Please select a CPL member.",

  // Validation documentaire
  "Valider et marquer le document": "Validate and stamp the document",
  "Choisissez le document validé. Un PDF portant la mention ci-dessous en haut à droite de chaque page sera enregistré dans le dossier de la demande.": "Choose the validated document. A PDF bearing the stamp below in the top right corner of each page will be saved in the request folder.",
  "Fichier": "File",
  "Dernière modification": "Last modified",
  "Date de validation": "Validation date",
  "Mention apposée": "Stamp applied",
  "Fichier créé": "File created",
  "Le PDF marqué est la version officielle validée. Le document source n'est pas modifié. Si le PDF source est signé électroniquement, sa signature ne sera plus valable sur la version marquée.": "The stamped PDF is the official validated version. The source document is not modified. If the source PDF is electronically signed, its signature will no longer be valid on the stamped version.",
  "Marquer aussi le fichier d'origine (PPTX ou DOCX, enregistré sous un nouveau nom, l'original n'est pas modifié). La mention est placée en haut à droite de chaque slide ou, pour un DOCX, dans les en-têtes existants.": "Also stamp the original file (PPTX or DOCX, saved under a new name, the original is not modified). The stamp is placed in the top right corner of each slide or, for a DOCX, in the existing headers.",
  "Annuler": "Cancel",
  "Téléchargement du document…": "Downloading the document…",
  "Conversion en PDF (SharePoint)…": "Converting to PDF (SharePoint)…",
  "Contrôle du PDF source…": "Checking the source PDF…",
  "Contrôle du PDF marqué…": "Checking the stamped PDF…",
  "Enregistrement du PDF marqué…": "Saving the stamped PDF…",
  "Passage au statut Traité…": "Moving to Processed status…",
  "Marquage du fichier d'origine…": "Stamping the original file…",
  "Contrôle du fichier d'origine marqué…": "Checking the stamped original file…",
  "Aucun document PPTX, DOCX ou PDF à valider dans le dossier de la demande.": "No PPTX, DOCX or PDF document to validate in the request folder.",

  // Dashboard KPI
  "Période": "Period",
  "Du": "From",
  "Au": "To",
  "30 derniers jours": "Last 30 days",
  "Trimestre en cours": "Current quarter",
  "Année en cours": "Current year",
  "12 derniers mois": "Last 12 months",
  "Depuis le début": "Since the beginning",
  "Dates personnalisées": "Custom dates",
  "Type de demande": "Request type",
  "(non attribué)": "(unassigned)",
  "Copier en image": "Copy as image",
  "Demandes reçues": "Requests received",
  "Demandes traitées": "Requests processed",
  "Stock en cours": "Open backlog",
  "(actuel)": "(current)",
  "Hors SLA en cours": "Open beyond SLA",
  "Respect du SLA": "SLA compliance",
  "Délai médian de traitement": "Median processing time",
  "Délai médian de prise en charge": "Median time to pick up",
  "soumission → premier statut Compliance": "submission → first Compliance status",
  "Taux de rejet": "Rejection rate",
  "(sur les reçues)": "(of received)",
  "période précédente : 0": "previous period: 0",
  "Reçues": "Received",
  "Traitées": "Processed",
  "En cours": "Open",
  "Hors SLA": "Beyond SLA",
  "Délai méd.": "Median time",
  "Aucune donnée sur la période.": "No data for this period.",
  "Aucune donnée": "No data",
  "Ancienneté du stock en cours": "Age of open backlog",
  "(jours ouvrés)": "(business days)",
  "0 à 2 j": "0 to 2 d",
  "3 à 5 j": "3 to 5 d",
  "6 à 10 j": "6 to 10 d",
  "Plus de 10 j": "Over 10 d",
  "Par type de demande": "By request type",
  "Par département": "By department",
  "Par responsable cpl": "By CPL owner",
  "Par priorité": "By priority",
  "Tracker CPL · Dashboard KPI": "Tracker CPL · KPI Dashboard",
  "Voir les demandes soumises ce mois-ci": "Show requests submitted this month",
  "Voir ces demandes dans la liste": "Show these requests in the list",
  "Évolution limitée aux 24 derniers mois.": "Trend limited to the last 24 months.",
  "Délais en jours ouvrés, hors attente d'informations. Stock, hors SLA et ancienneté : état actuel, indépendant de la période.": "Times in business days, excluding time awaiting information. Backlog, beyond SLA and age: current state, independent of the period.",
  "Filtres appliqués à la liste des demandes (la période n'est pas reprise).": "Filters applied to the request list (the period is not carried over).",
  "Capture indisponible sur cette page.": "Capture unavailable on this page.",
  "Dashboard copié dans le presse-papier, collez-le dans votre reporting.": "Dashboard copied to the clipboard, paste it into your report.",
  "Presse-papier indisponible : image téléchargée à la place.": "Clipboard unavailable: image downloaded instead.",

  // Export CSV du dashboard
  "Tracker CPL - Dashboard KPI": "Tracker CPL - KPI Dashboard",
  "Filtres": "Filters",
  "Aucun": "None",
  "Généré le": "Generated on",
  "Indicateur": "Indicator",
  "Valeur": "Value",
  "Période précédente": "Previous period",
  "Stock en cours (actuel)": "Open backlog (current)",
  "dont Réponse à revoir": "of which Reply to review",
  "dont En attente d'informations": "of which Awaiting information",
  "Respect du SLA (%)": "SLA compliance (%)",
  "Délai médian de traitement (j ouvrés)": "Median processing time (business days)",
  "Délai moyen de traitement (j ouvrés)": "Average processing time (business days)",
  "Délai médian de prise en charge (j ouvrés)": "Median time to pick up (business days)",
  "Taux de rejet (%)": "Rejection rate (%)",
  "Semaine du": "Week of",
  "Mois": "Month",
  "Ancienneté du stock": "Backlog age",
  "Délai médian (j ouvrés)": "Median time (business days)",
  "Respect SLA (%)": "SLA compliance (%)"
};

const MOIS_COURTS_EN = {
  'janv.': 'Jan', 'févr.': 'Feb', 'mars': 'Mar', 'avr.': 'Apr', 'mai': 'May', 'juin': 'Jun',
  'juil.': 'Jul', 'août': 'Aug', 'sept.': 'Sep', 'oct.': 'Oct', 'nov.': 'Nov', 'déc.': 'Dec'
};

// Fragment traduit par le dictionnaire ou les motifs, sinon inchangé
function trInner(s) { const t = trCore(s); return t == null ? s : t; }

// Valeur traduite si connue (statuts, types, libellés), sinon inchangée
function trv(v) { return TRACKER_EN[v] != null ? TRACKER_EN[v] : v; }
const plur = (n, s, p) => (Number(n) > 1 ? p : s);

// Libellé de période du dashboard ("du X au Y", "depuis le début", "jusqu'au X")
function trPeriodLabel(l) {
  let m;
  if ((m = l.match(/^du (\S+) au (\S+)$/))) return `from ${m[1]} to ${m[2]}`;
  if ((m = l.match(/^jusqu'au (\S+)$/))) return `until ${m[1]}`;
  if (l === 'depuis le début') return 'since the beginning';
  return l;
}

// Filtre "Libellé : valeur"
function trLabelValue(seg) {
  const m = seg.match(/^(.+?) : (.*)$/);
  return m ? `${trv(m[1])}: ${trv(m[2])}` : trv(seg);
}

const TRACKER_EN_PATTERNS = [
  // En-tête de la liste
  [/^(\d+) demandes au total · (\d+) ouvertes · (\d+) hors SLA( ·)?$/, (_, a, b, c, d) => `${a} requests in total · ${b} open · ${c} beyond SLA${d || ''}`],
  [/^(\d+) demandes? · (\d+) actions? requises? de votre part( ·)?$/, (_, a, b, d) => `${a} ${plur(a, 'request', 'requests')} · ${b} ${plur(b, 'action', 'actions')} required from you${d || ''}`],
  [/^Actualisé à (.+)$/, (_, a) => `Refreshed at ${a}`],
  [/^✕ Effacer les filtres \((\d+)\)$/, (_, a) => `✕ Clear filters (${a})`],

  // Nouvelle demande
  [/^Demandeur : (.+) · Référence générée à la soumission$/, (_, a) => `Requester: ${a} · Reference generated on submission`],
  [/^Upload des fichiers \((\d+)\/(\d+)\)\.\.\.$/, (_, a, b) => `Uploading files (${a}/${b})...`],
  [/^([\d.]+) (Mo|Ko)$/, (_, a, u) => `${a} ${u === 'Mo' ? 'MB' : 'KB'}`],
  [/^([\d.]+) Ko · (.*)$/, (_, a, b) => `${a} KB · ${b}`],
  [/^Demande issue de l'analyse Doc Review (\S+) : le fichier d'origine et le rapport de pré-contrôle seront joints automatiquement\. Complétez les autres champs puis soumettez\.$/,
    (_, a) => `Request created from Doc Review analysis ${a}: the original file and the pre-check report will be attached automatically. Complete the other fields, then submit.`],

  // Détail
  [/^par (.+)$/, (_, a) => `by ${a}`],
  [/^Analyse Doc Review liée : (.+)$/, (_, a) => `Linked Doc Review analysis: ${a}`],

  // Notifications
  [/^Erreur : (.*)$/, (_, a) => `Error: ${a}`],
  [/^Erreur lors de la capture : (.*)$/, (_, a) => `Capture error: ${a}`],
  [/^Demande (\S+) soumise avec succès\.$/, (_, a) => `Request ${a} submitted successfully.`],
  [/^Demande (\S+) créée, mais les fichiers de l'analyse (\S+) n'ont pas été joints : (.*)\. Utilisez « Finaliser la liaison Doc Review » dans la demande\.$/,
    (_, a, b, c) => `Request ${a} created, but the files of analysis ${b} were not attached: ${c}. Use "Finalize Doc Review link" in the request.`],
  [/^Soumission de l'analyse (\S+) impossible : (.*)$/, (_, a, b) => `Analysis ${a} cannot be submitted: ${b}`],
  [/^Analyse (\S+) non chargée : (.*)$/, (_, a, b) => `Analysis ${a} not loaded: ${b}`],
  [/^Fichiers de l'analyse (\S+) joints à la demande (\S+)\.$/, (_, a, b) => `Files of analysis ${a} attached to request ${b}.`],
  [/^Liaison impossible : (.*)$/, (_, a) => `Link failed: ${a}`],
  [/^Statut mis à jour : (.*)$/, (_, a) => `Status updated: ${trInner(a)}`],
  [/^Demande attribuée à (.*)\.$/, (_, a) => `Request assigned to ${a}.`],
  [/^Validation impossible : (.*)$/, (_, a) => `Validation not possible: ${a}`],
  [/^Document validé et marqué : (.*)\. Demande passée au statut Traité\.$/, (_, a) => `Document validated and stamped: ${a}. Request moved to Processed status.`],
  [/^Fichier d'origine marqué : (.*)$/, (_, a) => `Original file stamped: ${a}`],
  [/^Fichier d'origine : (.*)$/, (_, a) => `Original file: ${a}`],
  [/^Marquage des (\d+) page\(s\)…$/, (_, a) => `Stamping ${a} page(s)…`],
  [/^Validation interrompue, statut inchangé : (.*)$/, (_, a) => `Validation interrupted, status unchanged: ${a}`],

  // Dashboard
  [/^dont (\d+) réponses? à revoir, (\d+) en attente d'infos$/, (_, a, b) => `incl. ${a} ${plur(a, 'reply', 'replies')} to review, ${b} awaiting info`],
  [/^\(plus de (\d+) j ouvrés\)$/, (_, a) => `(over ${a} business days)`],
  [/^\(traitées en (\d+) j ouvrés max\.\)$/, (_, a) => `(processed within ${a} business days)`],
  [/^moyenne (.+), hors attente d'infos$/, (_, a) => `average ${trInner(a)}, excluding awaiting info`],
  [/^([▲▼=]) (\d+) % vs période précédente \((\d+)\)$/, (_, a, b, c) => `${a} ${b} % vs previous period (${c})`],
  [/^(\d+),(\d) j$/, (_, a, b) => `${a}.${b} d`],
  [/^Reçues et traitées par (semaine|mois)$/, (_, a) => `Received and processed per ${a === 'semaine' ? 'week' : 'month'}`],
  [/^(\d+) reçues?$/, (_, a) => `${a} received`],
  [/^(\d+) traitées?$/, (_, a) => `${a} processed`],
  [/^([a-zéû]+\.?) (\d{2})$/, (_, a, b) => (MOIS_COURTS_EN[a] ? `${MOIS_COURTS_EN[a]} ${b}` : null)],
  [/^· Période (.*)$/, (_, rest) => {
    const parts = rest.split(' · ');
    return '· Period ' + [trPeriodLabel(parts[0])].concat(parts.slice(1).map(trLabelValue)).join(' · ');
  }],
  [/^(\d+) demandes? dans le périmètre · généré le (\S+)$/, (_, a, b) => `${a} ${plur(a, 'request', 'requests')} in scope · generated on ${b}`],
  [/^(\d+) délais? estimés? faute d'historique complet \(date de dernière modification utilisée\)\.$/,
    (_, a) => `${a} ${plur(a, 'time', 'times')} estimated due to incomplete history (last modified date used).`],
  // Export CSV
  [/^Hors SLA en cours \(> (\d+) j ouvrés\)$/, (_, a) => `Open beyond SLA (> ${a} business days)`],
  [/^(du \S+ au \S+|jusqu'au \S+|depuis le début)$/, (_, a) => trPeriodLabel(a)],
  [/^(Type de demande|Département|Responsable CPL|Priorité) : (.*)$/, (s) => s.split(' / ').map(trLabelValue).join(' / ')]
];

let trackerLang = (() => { try { return localStorage.getItem('compliance_lang') === 'en' ? 'en' : 'fr'; } catch (e) { return 'fr'; } })();

function trCore(s) {
  if (TRACKER_EN[s] != null) return TRACKER_EN[s];
  for (const [re, fn] of TRACKER_EN_PATTERNS) {
    const m = s.match(re);
    if (m) { const t = fn(...m); if (t != null) return t; }   // null : motif non applicable
  }
  const m = s.match(/^(.+?) ?:$/);   // "Libellé :" (détail, récapitulatif de validation)
  if (m && TRACKER_EN[m[1]] != null) return TRACKER_EN[m[1]] + ':';
  return null;
}

// Traduit un texte de l'interface si la langue est l'anglais (inchangé sinon ou si inconnu)
function trText(s) {
  if (trackerLang !== 'en' || typeof s !== 'string') return s;
  const norm = s.replace(/[\s ]+/g, ' ').trim();
  if (!norm) return s;
  const t = trCore(norm);
  if (t == null || t === norm) return s;
  return (/^[\s ]/.test(s) ? ' ' : '') + t + (/[\s ]$/.test(s) ? ' ' : '');
}

// ── Traduction du DOM ──
const I18N_ATTRS = ['placeholder', 'title', 'data-placeholder', 'aria-label'];
const i18nText = new WeakMap();   // nœud texte → { fr, en }
const i18nAttr = new WeakMap();   // élément → { attr: { fr, en } }

function i18nSkip(el) {
  return !el || el.isContentEditable || !!el.closest('script,style,textarea,.no-i18n');
}

function i18nTranslateText(node) {
  if (i18nSkip(node.parentElement)) return;
  const cur = node.nodeValue;
  const rec = i18nText.get(node);
  if (rec && rec.en === cur) return;
  const en = trText(cur);
  if (en !== cur) { i18nText.set(node, { fr: cur, en }); node.nodeValue = en; }
}

// Les attributs (placeholder des zones de saisie comprises) sont traduits, sauf dans .no-i18n
function i18nTranslateAttrs(el) {
  if (el.closest('script,style,.no-i18n')) return;
  I18N_ATTRS.forEach(a => {
    if (!el.hasAttribute(a)) return;
    const cur = el.getAttribute(a);
    const recs = i18nAttr.get(el) || {};
    if (recs[a] && recs[a].en === cur) return;
    const en = trText(cur);
    if (en !== cur) { recs[a] = { fr: cur, en }; i18nAttr.set(el, recs); el.setAttribute(a, en); }
  });
}

function i18nWalk(root, onText, onEl) {
  if (root.nodeType === 3) { onText(root); return; }
  if (root.nodeType !== 1) return;
  onEl(root);
  const w = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  let n;
  while ((n = w.nextNode())) { if (n.nodeType === 3) onText(n); else onEl(n); }
}

function i18nTranslateTree(root) { i18nWalk(root, i18nTranslateText, i18nTranslateAttrs); }

function i18nRestoreTree(root) {
  i18nWalk(root,
    n => { const r = i18nText.get(n); if (r && n.nodeValue === r.en) n.nodeValue = r.fr; i18nText.delete(n); },
    el => {
      const recs = i18nAttr.get(el);
      if (!recs) return;
      Object.keys(recs).forEach(a => { if (el.getAttribute(a) === recs[a].en) el.setAttribute(a, recs[a].fr); });
      i18nAttr.delete(el);
    });
}

const i18nObserver = new MutationObserver(muts => {
  if (trackerLang !== 'en') return;
  muts.forEach(m => {
    if (m.type === 'childList') m.addedNodes.forEach(i18nTranslateTree);
    else if (m.type === 'characterData') i18nTranslateText(m.target);
    else if (m.type === 'attributes') i18nTranslateAttrs(m.target);
  });
});

function updateLangButton() {
  const btn = document.getElementById('lang-toggle-btn');
  if (btn) btn.textContent = trackerLang === 'fr' ? 'EN' : 'FR';
  document.documentElement.lang = trackerLang;
}

function setTrackerLang(lang) {
  trackerLang = lang === 'en' ? 'en' : 'fr';
  try { localStorage.setItem('compliance_lang', trackerLang); } catch (e) { /* préférence non mémorisée */ }
  if (trackerLang === 'en') i18nTranslateTree(document.body);
  else i18nRestoreTree(document.body);
  updateLangButton();
}

document.addEventListener('DOMContentLoaded', () => {
  i18nObserver.observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: I18N_ATTRS });
  if (trackerLang === 'en') i18nTranslateTree(document.body);
  updateLangButton();
});
