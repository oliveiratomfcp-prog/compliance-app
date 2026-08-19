// =============================================
//  TRANSLATIONS — FR / EN
//  Compliance App — Eiffel Investment Group
// =============================================

const TRANSLATIONS = {
  fr: {
    // NAVBAR
    nav_restricted: "Restricted List",
    nav_declare: "Déclarer",
    nav_history: "Historique",
    nav_role: "Collaborateur",

    // LOGIN
    login_title: "Compliance Portal",
    login_desc: "Connectez-vous avec votre compte d'entreprise Microsoft pour accéder à la plateforme.",
    login_btn: "Se connecter avec Microsoft",
    login_footer: "Accès réservé aux collaborateurs Eiffel IG",

    // RESTRICTED LIST
    restricted_title: "Restricted List",
    restricted_desc: "Liste des titres financiers sous restriction (information privilégiée ou NDA)",
    restricted_search: "Rechercher par nom ou code ISIN...",
    restricted_pdf_btn: "Générer l'attestation PDF",
    restricted_excel_btn: "Exporter Excel",
    restricted_bulk_btn: "Vérification en masse",
    restricted_loading: "Chargement de la liste...",
    restricted_error: "Impossible de charger la liste. Vérifiez vos permissions SharePoint.",
    restricted_realtime: "Mis à jour en temps réel depuis SharePoint",
    restricted_no_result: "Aucun résultat",
    col_company: "Société / Titre",
    col_isin: "Code ISIN",
    col_type: "Type de restriction",
    col_team: "Équipe",
    col_start: "Date de début",
    col_end: "Date de fin",

    // BULK CHECK
    bulk_title: "Vérification en masse",
    bulk_desc: "Collez vos codes ISIN ou noms de titres — un par ligne ou séparés par des virgules",
    bulk_placeholder: "FR0000131104\nFR0000120271\nApple\nTotalEnergies\n...",
    bulk_run_btn: "Lancer la vérification",
    bulk_restricted: "Restreint",
    bulk_clean: "Non restreint",
    bulk_summary: "titre(s) vérifié(s)",
    bulk_pdf_btn: "Rapport PDF",

    // DECLARE
    declare_title: "Déclarer",
    declare_desc: "Saisie Compliance et déclaration de transaction personnelle",
    nda_title: "NDA List",
    nda_desc: "Ajouter un nouveau titre sous NDA dans la liste de restriction. Ce formulaire alimentera directement la liste SharePoint NDA List.",
    nda_note: "Mise à jour immédiate de la Restricted List",
    nda_btn: "Ouvrir le formulaire NDA",
    infoprev_title: "Information privilégiée",
    infoprev_desc: "Ajouter un nouveau titre sous information privilégiée. Ce formulaire alimentera directement la liste SharePoint Information privilégiée.",
    infoprev_note: "Mise à jour immédiate de la Restricted List",
    infoprev_btn: "Ouvrir le formulaire Info privilégiée",
    transaction_title: "Demande de transaction personnelle",
    transaction_desc: "Soumettez une demande de transaction sur un titre financier et générez votre attestation PDF à envoyer à CPL sous 24h.",
    transaction_click: "Cliquez pour ouvrir le formulaire",
    transaction_btn: "Déclarer une transaction",

    // FORM
    form_title: "Demande de transaction personnelle",
    form_subtitle: "Remplissez le formulaire et générez votre attestation PDF",
    form_alert: "La demande doit être faite <strong>avant tout passage d'ordre</strong> et à joindre au mail généré automatiquement.",
    form_titre: "Nom du titre financier *",
    form_titre_placeholder: "Ex: Apple Inc., Total SA...",
    form_type: "Type d'opération *",
    form_type_placeholder: "-- Sélectionner --",
    form_type_buy: "Achat",
    form_type_sell: "Vente",
    form_qty: "Quantité *",
    form_qty_placeholder: "Ex: 100",
    form_price: "Prix unitaire (€)",
    form_price_placeholder: "Ex: 145.50",
    form_date: "Date de transaction *",
    form_account: "Compte utilisé *",
    form_account_placeholder: "Ex: Compte-titres BNP, PEA Boursorama...",
    form_comment: "Commentaire (optionnel)",
    form_comment_placeholder: "Informations complémentaires...",
    form_submit_btn: "Générer attestation et mail",
    form_success: "✓ PDF généré ! Votre client mail va s'ouvrir pour envoyer l'attestation à CPL.",
    form_error_required: "Veuillez remplir tous les champs obligatoires (*)",
    form_error_restricted: "figure sur la Restricted List — cette transaction ne peut pas être demandée. Contactez le service Compliance.",

    // RULES TOOLTIP
    rules_title: "Règles — Transactions personnelles",
    rules_btn: "Règles applicables",
    rules: [
      "🚫 <strong>Interdit depuis le 1er avril 2026 :</strong> toute transaction personnelle sur titres vifs cotés.",
      "⚠️ <strong>Cession de titres déjà en portefeuille :</strong> accord préalable du RCCI obligatoire via \"demande de transaction personnelle\".",
      "✅ <strong>Autorisé :</strong> fonds gérés par Eiffel IG, fonds non gérés, trackers, puts, calls, warrants sur indices cotés.",
      "📋 <strong>Déclaration obligatoire :</strong> transactions sur fonds Eiffel IG ou vente de titre déjà en PTF + envoi de l'avis d'opéré sous 48h.",
      "📅 <strong>Annuellement :</strong> remise au RCCI du relevé complet des opérations sur instruments financiers."
    ],

    // HISTORY
    history_title: "Mon historique",
    history_desc: "Vos déclarations et consultations passées",
    history_refresh: "Rafraîchir",
    history_search: "Rechercher par nom de titre...",
    history_empty: "Aucune déclaration pour l'instant",
    history_loading: "Chargement de l'historique...",
    col_h_titre: "Titre",
    col_h_type: "Type",
    col_h_qty: "Quantité / Statut",
    col_h_price: "Prix",
    col_h_date_tx: "Date transaction",
    col_h_date_decl: "Date",
    col_h_pdf: "Attestation",
    badge_consultation: "Consultation",
    badge_declaration: "Déclaration",
    badge_restricted: "Restreint",
    badge_clean: "Non restreint",

    // GIFT FORM (Registre cadeaux)
    gift_alert: "Il n'existe aucun seuil de déclaration : tout cadeau ou invitation reçu(e) ou offert(e) doit être déclaré(e), quel que soit son montant.",
    gift_intro: "Avant d'accepter ou d'offrir un cadeau ou une invitation, il convient de s'assurer que sa valeur reste raisonnable au regard du niveau de vie local, qu'il s'inscrit dans un contexte professionnel, et qu'il ne risque pas d'influencer une décision. Demandez-vous également si vous seriez capable d'en parler librement à votre manager ? Si vous hésitez à le mentionner spontanément, c'est probablement le signe qu'il pose problème.<br><br>En cas de doute ou de réponse négative à l'un de ces points, il est recommandé de refuser le cadeau ou l'invitation et de solliciter l'avis de votre manager ou du département Conformité.",
    gift_label_type_cadeau: "Type de cadeau *",
    gift_opt_select: "-- Sélectionner --",
    gift_opt_type: ["Cadeau matériel", "Repas", "Invitation sportive", "Invitation culturelle", "Séminaire", "Conférence", "Formation", "Déplacement", "Hébergement", "Remise commerciale", "Autre"],
    gift_label_nature_tiers: "Nature du tiers *",
    gift_opt_nature_tiers: ["Investisseur", "Prospect investisseur", "Société en portefeuille", "Cible d'investissement", "Dépositaire/Valorisateur", "Banque d'affaires", "Contrepartie/Broker", "Autre Prestataire", "Fournisseur", "Distributeur", "Conseil juridique/fiscal", "Autre"],
    gift_label_nature_tiers_precision: "Précisez la nature du tiers *",
    gift_label_operation_en_cours: "Une opération, négociation, levée de fonds, due diligence ou renouvellement de contrat est-elle en cours avec ce tiers ? *",
    gift_radio_oui: "Oui",
    gift_radio_non: "Non",
    gift_label_date: "Date à laquelle le cadeau a été reçu/offert *",
    gift_label_sort_cadeau: "Sort du cadeau *",
    gift_opt_sort_cadeau: ["Conservé à titre personnel", "Partagé avec l'équipe", "Mis à disposition de l'ensemble des collaborateurs", "Utilisé dans le cadre professionnel", "Refusé", "Autre"],
    gift_label_sort_cadeau_precision: "Précisez le sort du cadeau *",
  },

  en: {
    // NAVBAR
    nav_restricted: "Restricted List",
    nav_declare: "Declare",
    nav_history: "History",
    nav_role: "Employee",

    // LOGIN
    login_title: "Compliance Portal",
    login_desc: "Sign in with your Microsoft corporate account to access the platform.",
    login_btn: "Sign in with Microsoft",
    login_footer: "Access restricted to Eiffel IG employees",

    // RESTRICTED LIST
    restricted_title: "Restricted List",
    restricted_desc: "List of financial securities under restriction (privileged information or NDA)",
    restricted_search: "Search by name or ISIN code...",
    restricted_pdf_btn: "Generate PDF certificate",
    restricted_excel_btn: "Export Excel",
    restricted_bulk_btn: "Bulk check",
    restricted_loading: "Loading list...",
    restricted_error: "Unable to load list. Please check your SharePoint permissions.",
    restricted_realtime: "Updated in real time from SharePoint",
    restricted_no_result: "No results",
    col_company: "Company / Security",
    col_isin: "ISIN Code",
    col_type: "Restriction type",
    col_team: "Team",
    col_start: "Start date",
    col_end: "End date",

    // BULK CHECK
    bulk_title: "Bulk check",
    bulk_desc: "Paste your ISIN codes or security names — one per line or separated by commas",
    bulk_placeholder: "FR0000131104\nFR0000120271\nApple\nTotalEnergies\n...",
    bulk_run_btn: "Run check",
    bulk_restricted: "Restricted",
    bulk_clean: "Approved",
    bulk_summary: "securities checked",
    bulk_pdf_btn: "PDF report",

    // DECLARE
    declare_title: "Declare",
    declare_desc: "Compliance entry and personal transaction declaration",
    nda_title: "NDA List",
    nda_desc: "Add a new title under NDA to the restriction list. This form will directly update the SharePoint NDA List.",
    nda_note: "Immediate update of the Restricted List",
    nda_btn: "Open NDA form",
    infoprev_title: "Privileged information",
    infoprev_desc: "Add a new title under privileged information. This form will directly update the SharePoint list.",
    infoprev_note: "Immediate update of the Restricted List",
    infoprev_btn: "Open privileged info form",
    transaction_title: "Personal transaction request",
    transaction_desc: "Submit a transaction request on a financial security and generate your PDF certificate to send to CPL within 24h.",
    transaction_click: "Click to open the form",
    transaction_btn: "Declare a transaction",

    // FORM
    form_title: "Personal transaction request",
    form_subtitle: "Fill in the form and generate your PDF certificate",
    form_alert: "The request must be submitted <strong>before any order is placed</strong> and attached to the automatically generated email.",
    form_titre: "Security name *",
    form_titre_placeholder: "Ex: Apple Inc., Total SA...",
    form_type: "Transaction type *",
    form_type_placeholder: "-- Select --",
    form_type_buy: "Buy",
    form_type_sell: "Sell",
    form_qty: "Quantity *",
    form_qty_placeholder: "Ex: 100",
    form_price: "Unit price (€)",
    form_price_placeholder: "Ex: 145.50",
    form_date: "Transaction date *",
    form_account: "Account used *",
    form_account_placeholder: "Ex: BNP securities account, Boursorama PEA...",
    form_comment: "Comment (optional)",
    form_comment_placeholder: "Additional information...",
    form_submit_btn: "Generate certificate & email",
    form_success: "✓ PDF generated! Your email client will open to send the certificate to CPL.",
    form_error_required: "Please fill in all required fields (*)",
    form_error_restricted: "is on the Restricted List — this transaction cannot be requested. Contact the Compliance team.",

    // RULES TOOLTIP
    rules_title: "Rules — Personal transactions",
    rules_btn: "Applicable rules",
    rules: [
      "🚫 <strong>Prohibited from April 1st, 2026:</strong> all personal transactions on listed equities.",
      "⚠️ <strong>Disposal of existing portfolio holdings:</strong> prior approval from the RCCI required via \"personal transaction request\".",
      "✅ <strong>Authorised:</strong> Eiffel IG-managed funds, non-managed funds, trackers, puts, calls, warrants on regulated market indices.",
      "📋 <strong>Mandatory declaration:</strong> transactions on Eiffel IG funds or disposal of existing portfolio holdings + confirmation notice within 48h.",
      "📅 <strong>Annually:</strong> submission to RCCI of full record of transactions on financial instrument accounts."
    ],

    // HISTORY
    history_title: "My history",
    history_desc: "Your past declarations and consultations",
    history_refresh: "Refresh",
    history_search: "Search by security name...",
    history_empty: "No declarations yet",
    history_loading: "Loading history...",
    col_h_titre: "Security",
    col_h_type: "Type",
    col_h_qty: "Quantity / Status",
    col_h_price: "Price",
    col_h_date_tx: "Transaction date",
    col_h_date_decl: "Date",
    col_h_pdf: "Certificate",
    badge_consultation: "Consultation",
    badge_declaration: "Declaration",
    badge_restricted: "Restricted",
    badge_clean: "Approved",

    // GIFT FORM (Gift register)
    gift_alert: "There is no reporting threshold: any gift or invitation received or given must be declared, regardless of its value.",
    gift_intro: "Before accepting or offering a gift or invitation, make sure its value remains reasonable given the local standard of living, that it takes place in a professional context, and that it is not likely to influence a decision. Also ask yourself whether you would be comfortable talking about it openly with your manager. If you hesitate to mention it spontaneously, that is probably a sign that it is a problem.<br><br>In case of doubt or a negative answer to any of these points, it is recommended to decline the gift or invitation and to seek the advice of your manager or the Compliance department.",
    gift_label_type_cadeau: "Type of gift *",
    gift_opt_select: "-- Select --",
    gift_opt_type: ["Material gift", "Meal", "Sports invitation", "Cultural invitation", "Seminar", "Conference", "Training", "Travel", "Accommodation", "Commercial discount", "Other"],
    gift_label_nature_tiers: "Nature of the third party *",
    gift_opt_nature_tiers: ["Investor", "Prospective investor", "Portfolio company", "Investment target", "Depositary/Valuation agent", "Investment bank", "Counterparty/Broker", "Other service provider", "Supplier", "Distributor", "Legal/tax adviser", "Other"],
    gift_label_nature_tiers_precision: "Specify the nature of the third party *",
    gift_label_operation_en_cours: "Is a transaction, negotiation, fundraising, due diligence or contract renewal currently in progress with this third party? *",
    gift_radio_oui: "Yes",
    gift_radio_non: "No",
    gift_label_date: "Date the gift was received/given *",
    gift_label_sort_cadeau: "Outcome of the gift *",
    gift_opt_sort_cadeau: ["Kept for personal use", "Shared with the team", "Made available to all employees", "Used for professional purposes", "Declined", "Other"],
    gift_label_sort_cadeau_precision: "Specify the outcome of the gift *",
  }
};

let currentLang = "fr";

function t(key) {
  return TRANSLATIONS[currentLang][key] || TRANSLATIONS["fr"][key] || key;
}

function setLang(lang) {
  currentLang = lang;
  applyTranslations();
}

function applyTranslations() {
  // Navbar - on utilise les spans avec IDs
  setText("nav-restricted-el", t("nav_restricted"));
  setText("nav-declare-el", t("nav_declare"));
  setText("nav-history-el", t("nav_history"));
  setText("nav-role-el", t("nav_role"));

  // Login
  setText("login-title-el", t("login_title"));
  setText("login-desc-el", t("login_desc"));
  setText("login-btn-el", t("login_btn"));
  setText("login-footer-el", t("login_footer"));

  // Restricted List
  setText("restricted-title-el", t("restricted_title"));
  setText("restricted-desc-el", t("restricted_desc"));
  setPlaceholder("search-restricted", t("restricted_search"));
  setText("restricted-pdf-btn-el", t("restricted_pdf_btn"));
  setText("restricted-excel-btn-el", t("restricted_excel_btn"));
  setText("restricted-bulk-btn-el", t("restricted_bulk_btn"));
  setText("restricted-realtime-el", t("restricted_realtime"));
  setText("col-company-el", t("col_company"));
  setText("col-isin-el", t("col_isin"));
  setText("col-type-el", t("col_type"));
  setText("col-team-el", t("col_team"));
  setText("col-start-el", t("col_start"));
  setText("col-end-el", t("col_end"));

  // Bulk
  setText("bulk-title-el", t("bulk_title"));
  setText("bulk-desc-el", t("bulk_desc"));
  setPlaceholder("bulk-input", t("bulk_placeholder"));
  setText("bulk-run-btn-el", t("bulk_run_btn"));

  // Declare
  setText("declare-title-el", t("declare_title"));
  setText("declare-desc-el", t("declare_desc"));
  setText("nda-title-el", t("nda_title"));
  setText("nda-desc-el", t("nda_desc"));
  setText("nda-note-el", t("nda_note"));
  setText("nda-btn-el", t("nda_btn"));
  setText("infoprev-title-el", t("infoprev_title"));
  setText("infoprev-desc-el", t("infoprev_desc"));
  setText("infoprev-note-el", t("infoprev_note"));
  setText("infoprev-btn-el", t("infoprev_btn"));
  setText("transaction-card-title-el", t("transaction_title"));
  setText("transaction-card-desc-el", t("transaction_desc"));
  setText("transaction-card-note-el", t("transaction_click"));
  setText("transaction-card-btn-el", t("transaction_btn"));

  // Form
  setText("form-title-el", t("form_title"));
  setText("form-subtitle-el", t("form_subtitle"));
  setHTML("form-alert-el", t("form_alert"));
  setText("form-label-titre", t("form_titre"));
  setPlaceholder("f-titre", t("form_titre_placeholder"));
  setText("form-label-type", t("form_type"));
  setOptionText("f-type", 0, t("form_type_placeholder"));
  setOptionText("f-type", 1, t("form_type_buy"));
  setOptionText("f-type", 2, t("form_type_sell"));
  setText("form-label-qty", t("form_qty"));
  setPlaceholder("f-quantite", t("form_qty_placeholder"));
  setText("form-label-price", t("form_price"));
  setPlaceholder("f-prix", t("form_price_placeholder"));
  setText("form-label-date", t("form_date"));
  setText("form-label-account", t("form_account"));
  setPlaceholder("f-compte", t("form_account_placeholder"));
  setText("form-label-comment", t("form_comment"));
  setPlaceholder("f-commentaire", t("form_comment_placeholder"));
  setText("form-submit-btn-el", t("form_submit_btn"));

  // Rules tooltip
  setText("rules-btn-el", t("rules_btn"));
  setText("rules-title-el", t("rules_title"));
  const rulesList = document.getElementById("rules-list-el");
  if (rulesList) rulesList.innerHTML = TRANSLATIONS[currentLang].rules.map(r => `<p>${r}</p>`).join("");

  // History
  setText("history-title-el", t("history_title"));
  setText("history-desc-el", t("history_desc"));
  setText("history-refresh-el", t("history_refresh"));
  setPlaceholder("search-history", t("history_search"));
  setText("history-empty-el", t("history_empty"));
  setText("col-h-titre-el", t("col_h_titre"));
  setText("col-h-type-el", t("col_h_type"));
  setText("col-h-qty-el", t("col_h_qty"));
  setText("col-h-price-el", t("col_h_price"));
  setText("col-h-date-tx-el", t("col_h_date_tx"));
  setText("col-h-date-decl-el", t("col_h_date_decl"));
  setText("col-h-pdf-el", t("col_h_pdf"));

  // Gift form (Registre cadeaux)
  setText("gift-alert-el", t("gift_alert"));
  setHTML("gift-intro-el", t("gift_intro"));
  setText("form-label-type-cadeau", t("gift_label_type_cadeau"));
  setOptionText("g-type-cadeau", 0, t("gift_opt_select"));
  TRANSLATIONS[currentLang].gift_opt_type.forEach((label, i) => setOptionText("g-type-cadeau", i + 1, label));
  setText("form-label-nature-tiers", t("gift_label_nature_tiers"));
  setOptionText("g-nature-tiers", 0, t("gift_opt_select"));
  TRANSLATIONS[currentLang].gift_opt_nature_tiers.forEach((label, i) => setOptionText("g-nature-tiers", i + 1, label));
  setText("form-label-nature-tiers-precision", t("gift_label_nature_tiers_precision"));
  setText("form-label-operation-en-cours", t("gift_label_operation_en_cours"));
  setText("radio-operation-oui-el", t("gift_radio_oui"));
  setText("radio-operation-non-el", t("gift_radio_non"));
  setText("form-label-date-gift", t("gift_label_date"));
  setText("form-label-sort-cadeau", t("gift_label_sort_cadeau"));
  setOptionText("g-sort-cadeau", 0, t("gift_opt_select"));
  TRANSLATIONS[currentLang].gift_opt_sort_cadeau.forEach((label, i) => setOptionText("g-sort-cadeau", i + 1, label));
  setText("form-label-sort-cadeau-precision", t("gift_label_sort_cadeau_precision"));

  // Lang button
  const btn = document.getElementById("lang-toggle-btn");
  if (btn) btn.textContent = currentLang === "fr" ? "EN" : "FR";

  // html lang attribute
  document.documentElement.lang = currentLang;
}

// Helpers
function setText(id, text) { const el = document.getElementById(id); if (el) el.textContent = text; }
function setHTML(id, html) { const el = document.getElementById(id); if (el) el.innerHTML = html; }
function setPlaceholder(id, text) { const el = document.getElementById(id); if (el) el.placeholder = text; }
function setOptionText(selectId, index, text) {
  const el = document.getElementById(selectId);
  if (el && el.options[index]) el.options[index].text = text;
}
