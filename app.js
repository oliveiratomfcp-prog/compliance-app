// =============================================
//  COMPLIANCE APP — EIFFEL INVESTMENT GROUP
//  Logique principale (authentification, 
//  SharePoint, génération PDF)
// =============================================

// -----------------------------------------------
// 1. CONFIGURATION
//    Tes identifiants Azure AD et SharePoint
// -----------------------------------------------
const CONFIG = {
  clientId: "56c83db2-7fc6-4826-8073-82077e91ae27",   // ID de ton application Azure AD
  tenantId: "eb5288b6-5ccb-4f5d-88fb-346db1317a4f",   // ID de ton organisation
  redirectUri: window.location.origin,                  // URL de redirection (automatique)
  sharepoint: {
    siteHostname: "eiffelinvestmentgroup.sharepoint.com",
    sitePath: "/sites/CPL",
    listRestricted: "NDA List",                         // Nom de ta liste Restricted
    listDeclarations: "Information privilégiée"         // Nom de ta liste Déclarations
  }
};

// -----------------------------------------------
// 2. INITIALISATION MSAL (Authentification Microsoft)
//    MSAL = Microsoft Authentication Library
//    C'est la bibliothèque qui gère le login
// -----------------------------------------------
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri
  },
  cache: {
    cacheLocation: "sessionStorage" // Stocke la session dans le navigateur
  }
};

// On crée l'instance MSAL
const msalInstance = new msal.PublicClientApplication(msalConfig);

// Les permissions qu'on demande à Microsoft Graph
const graphScopes = {
  scopes: [
    "User.Read",
    "Sites.Read.All",
    "Sites.ReadWrite.All"
  ]
};

// -----------------------------------------------
// 3. VARIABLES GLOBALES
// -----------------------------------------------
let currentUser = null;       // Infos de l'utilisateur connecté
let siteId = null;            // ID du site SharePoint (récupéré au démarrage)
let restrictedItems = [];     // Les titres de la restricted list

// -----------------------------------------------
// 4. DÉMARRAGE DE L'APPLICATION
// -----------------------------------------------
async function init() {
  try {
    // Gère le retour après connexion Microsoft
    await msalInstance.handleRedirectPromise();

    // Vérifie si un utilisateur est déjà connecté
    const accounts = msalInstance.getAllAccounts();

    if (accounts.length > 0) {
      // Utilisateur déjà connecté → on va sur l'app
      msalInstance.setActiveAccount(accounts[0]);
      await onLoggedIn();
    } else {
      // Personne connecté → on affiche l'écran de login
      showScreen("login");
    }
  } catch (err) {
    console.error("Erreur d'initialisation:", err);
    showScreen("login");
  }
}

// -----------------------------------------------
// 5. CONNEXION / DÉCONNEXION
// -----------------------------------------------

// Quand l'utilisateur clique "Se connecter avec Microsoft"
document.getElementById("btn-login").addEventListener("click", async () => {
  try {
    // Ouvre la fenêtre de connexion Microsoft en popup
    const result = await msalInstance.loginPopup(graphScopes);
    msalInstance.setActiveAccount(result.account);
    await onLoggedIn();
  } catch (err) {
    console.error("Erreur de connexion:", err);
    alert("Erreur de connexion. Vérifiez votre compte et réessayez.");
  }
});

// Quand l'utilisateur clique "Se déconnecter"
document.getElementById("btn-logout").addEventListener("click", () => {
  msalInstance.logoutPopup().then(() => {
    showScreen("login");
    currentUser = null;
  });
});

// Après connexion réussie
async function onLoggedIn() {
  // Récupère le profil de l'utilisateur connecté
  currentUser = await callGraphAPI("/me");

  // Affiche le nom et les initiales dans la navbar
  const fullName = currentUser.displayName || currentUser.mail;
  document.getElementById("user-name").textContent = fullName;
  const initials = fullName.split(" ").map(w => w[0]).join("").substring(0, 2).toUpperCase();
  document.getElementById("user-avatar").textContent = initials;

  // Récupère l'ID du site SharePoint (nécessaire pour toutes les requêtes)
  try {
    const site = await callGraphAPI(
      `/sites/${CONFIG.sharepoint.siteHostname}:${CONFIG.sharepoint.sitePath}`
    );
    siteId = site.id;
  } catch (err) {
    console.error("Impossible de trouver le site SharePoint:", err);
    alert("Erreur : impossible de se connecter au site SharePoint. Vérifiez les permissions.");
    return;
  }

  // Affiche l'application
  showScreen("app");

  // Charge les données du premier onglet
  loadRestrictedList();
}

// -----------------------------------------------
// 6. NAVIGATION ENTRE ONGLETS
// -----------------------------------------------
document.querySelectorAll(".nav-tab").forEach(tab => {
  tab.addEventListener("click", () => {
    const target = tab.dataset.tab;

    // Met à jour le style des onglets
    document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
    tab.classList.add("active");

    // Affiche la bonne section
    document.querySelectorAll(".tab-content").forEach(s => s.classList.remove("active"));
    document.getElementById(`tab-${target}`).classList.add("active");

    // Charge les données si nécessaire
    if (target === "history") loadHistory();
  });
});

// -----------------------------------------------
// 7. ONGLET 1 — RESTRICTED LIST
// -----------------------------------------------
async function loadRestrictedList() {
  showElement("restricted-loading");
  hideElement("restricted-table-container");
  hideElement("restricted-error");

  try {
    // Appel Graph API pour lire la liste SharePoint
    const data = await callGraphAPI(
      `/sites/${siteId}/lists/${encodeURIComponent(CONFIG.sharepoint.listRestricted)}/items?expand=fields`
    );

    restrictedItems = data.value || [];

    // Met à jour le compteur
    document.getElementById("restricted-count").textContent =
      `${restrictedItems.length} titre${restrictedItems.length > 1 ? "s" : ""}`;

    // Remplit le tableau
    const tbody = document.getElementById("restricted-tbody");
    tbody.innerHTML = "";

    if (restrictedItems.length === 0) {
      tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;color:#9aaaba;padding:32px">Aucun titre dans la liste</td></tr>`;
    } else {
      restrictedItems.forEach(item => {
        const f = item.fields || {};
        // Adapte ces noms de colonnes aux noms réels de ta liste SharePoint
        const titre = f.Title || f.Titre || f.Emetteur || "—";
        const typeRestriction = f.TypeRestriction || f.Type || f.Restriction || "Restriction";
        const dateAjout = f.Created ? new Date(f.Created).toLocaleDateString("fr-FR") : "—";
        const statut = f.Statut || "Actif";

        const badgeClass = typeRestriction.toLowerCase().includes("nda") ? "badge-nda" : "badge-restricted";

        tbody.innerHTML += `
          <tr>
            <td><strong>${titre}</strong></td>
            <td><span class="badge ${badgeClass}">${typeRestriction}</span></td>
            <td>${dateAjout}</td>
            <td><span class="badge badge-active">${statut}</span></td>
          </tr>`;
      });
    }

    hideElement("restricted-loading");
    showElement("restricted-table-container");

  } catch (err) {
    console.error("Erreur chargement restricted list:", err);
    hideElement("restricted-loading");
    showElement("restricted-error");
    document.getElementById("restricted-error-msg").textContent =
      "Impossible de charger la liste. Vérifiez vos permissions SharePoint.";
  }
}

// Bouton "Générer l'attestation PDF" pour la Restricted List
document.getElementById("btn-generate-restricted-pdf").addEventListener("click", () => {
  if (!currentUser) return;
  generateRestrictedPDF();
});

// -----------------------------------------------
// 8. ONGLET 2 — DÉCLARATION DE TRANSACTION
// -----------------------------------------------
document.getElementById("btn-submit-declaration").addEventListener("click", async () => {
  // Récupère les valeurs du formulaire
  const titre = document.getElementById("f-titre").value.trim();
  const type = document.getElementById("f-type").value;
  const quantite = document.getElementById("f-quantite").value;
  const prix = document.getElementById("f-prix").value;
  const date = document.getElementById("f-date").value;
  const compte = document.getElementById("f-compte").value.trim();
  const commentaire = document.getElementById("f-commentaire").value.trim();

  // Validation des champs obligatoires
  if (!titre || !type || !quantite || !date || !compte) {
    showFormMessage("Veuillez remplir tous les champs obligatoires (*)", "error");
    return;
  }

  const btn = document.getElementById("btn-submit-declaration");
  btn.disabled = true;
  btn.textContent = "Enregistrement...";

  try {
    // Prépare les données à envoyer dans SharePoint
    // ⚠️ Les noms de champs (Title, TypeOperation, etc.) doivent correspondre
    //    aux noms de colonnes exacts de ta liste SharePoint "Information privilégiée"
    const fields = {
      Title: titre,
      TypeOperation: type,
      Quantite: parseInt(quantite),
      PrixUnitaire: prix ? parseFloat(prix) : null,
      DateTransaction: date,
      CompteUtilise: compte,
      Commentaire: commentaire,
      Declarant: currentUser.displayName,
      DeclarantEmail: currentUser.mail,
      DateDeclaration: new Date().toISOString()
    };

    // Envoie dans SharePoint via Graph API
    await callGraphAPI(
      `/sites/${siteId}/lists/${encodeURIComponent(CONFIG.sharepoint.listDeclarations)}/items`,
      "POST",
      { fields }
    );

    showFormMessage("✓ Déclaration enregistrée avec succès !", "success");

    // Génère automatiquement le PDF de la déclaration
    const declarationData = { titre, type, quantite, prix, date, compte, commentaire };
    generateDeclarationPDF(declarationData);

    // Remet le formulaire à zéro
    setTimeout(() => {
      document.getElementById("f-titre").value = "";
      document.getElementById("f-type").value = "";
      document.getElementById("f-quantite").value = "";
      document.getElementById("f-prix").value = "";
      document.getElementById("f-date").value = "";
      document.getElementById("f-compte").value = "";
      document.getElementById("f-commentaire").value = "";
      hideElement("declaration-msg");
    }, 4000);

  } catch (err) {
    console.error("Erreur enregistrement déclaration:", err);
    showFormMessage("Erreur lors de l'enregistrement. Vérifiez vos permissions.", "error");
  } finally {
    btn.disabled = false;
    btn.innerHTML = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg> Enregistrer la déclaration`;
  }
});

// -----------------------------------------------
// 9. ONGLET 3 — HISTORIQUE
// -----------------------------------------------
async function loadHistory() {
  showElement("history-loading");
  hideElement("history-table-container");
  hideElement("history-empty");

  try {
    // Récupère uniquement les déclarations de l'utilisateur connecté
    // On filtre par email du déclarant
    const email = currentUser.mail;
    const data = await callGraphAPI(
      `/sites/${siteId}/lists/${encodeURIComponent(CONFIG.sharepoint.listDeclarations)}/items?expand=fields&$filter=fields/DeclarantEmail eq '${email}'`
    );

    const items = data.value || [];
    hideElement("history-loading");

    if (items.length === 0) {
      showElement("history-empty");
      return;
    }

    // Remplit le tableau d'historique
    const tbody = document.getElementById("history-tbody");
    tbody.innerHTML = "";

    items.forEach(item => {
      const f = item.fields || {};
      const dateDeclaration = f.DateDeclaration
        ? new Date(f.DateDeclaration).toLocaleDateString("fr-FR")
        : new Date(item.createdDateTime).toLocaleDateString("fr-FR");

      tbody.innerHTML += `
        <tr>
          <td><strong>${f.Title || "—"}</strong></td>
          <td>${f.TypeOperation || "—"}</td>
          <td>${f.Quantite || "—"}</td>
          <td>${f.PrixUnitaire ? f.PrixUnitaire + " €" : "—"}</td>
          <td>${f.DateTransaction ? new Date(f.DateTransaction).toLocaleDateString("fr-FR") : "—"}</td>
          <td>${dateDeclaration}</td>
          <td>
            <button class="btn-pdf" onclick="regeneratePDF(${JSON.stringify(f).replace(/"/g, '&quot;')})">
              <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              PDF
            </button>
          </td>
        </tr>`;
    });

    showElement("history-table-container");

  } catch (err) {
    console.error("Erreur chargement historique:", err);
    hideElement("history-loading");
    showElement("history-empty");
  }
}

// Regénère un PDF depuis l'historique
function regeneratePDF(fields) {
  const data = {
    titre: fields.Title,
    type: fields.TypeOperation,
    quantite: fields.Quantite,
    prix: fields.PrixUnitaire,
    date: fields.DateTransaction,
    compte: fields.CompteUtilise,
    commentaire: fields.Commentaire
  };
  generateDeclarationPDF(data);
}

// -----------------------------------------------
// 10. GÉNÉRATION PDF — RESTRICTED LIST
// -----------------------------------------------
function generateRestrictedPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  const now = new Date();
  const dateStr = now.toLocaleDateString("fr-FR", {
    weekday: "long", year: "numeric", month: "long", day: "numeric"
  });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;

  // --- En-tête ---
  doc.setFillColor(13, 27, 42);          // Bleu marine
  doc.rect(0, 0, 210, 40, "F");

  doc.setTextColor(201, 168, 76);        // Or
  doc.setFontSize(20);
  doc.setFont("helvetica", "bold");
  doc.text("EIFFEL INVESTMENT GROUP", 20, 18);

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.text("Compliance Portal", 20, 28);

  // --- Titre du document ---
  doc.setTextColor(13, 27, 42);
  doc.setFontSize(16);
  doc.setFont("helvetica", "bold");
  doc.text("ATTESTATION DE CONSULTATION", 105, 58, { align: "center" });

  doc.setFontSize(12);
  doc.setFont("helvetica", "normal");
  doc.text("Restricted List — Titres sous restriction", 105, 68, { align: "center" });

  // --- Ligne de séparation ---
  doc.setDrawColor(201, 168, 76);
  doc.setLineWidth(0.8);
  doc.line(20, 74, 190, 74);

  // --- Informations de l'attestation ---
  doc.setFontSize(11);
  doc.setTextColor(50, 50, 50);

  let y = 90;
  const addLine = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.setTextColor(13, 27, 42);
    doc.text(label, 20, y);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(50, 50, 50);
    doc.text(value, 80, y);
    y += 10;
  };

  addLine("Collaborateur :", userName);
  addLine("Email :", currentUser.mail || "—");
  addLine("Date de consultation :", dateStr);
  addLine("Heure de consultation :", timeStr);
  addLine("Nombre de titres consultés :", `${restrictedItems.length} titre(s)`);

  // --- Texte d'attestation ---
  y += 10;
  doc.setDrawColor(220, 220, 220);
  doc.setLineWidth(0.3);
  doc.rect(20, y - 5, 170, 50);

  doc.setFillColor(248, 249, 250);
  doc.rect(20, y - 5, 170, 50, "F");

  doc.setFontSize(10.5);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(50, 50, 50);

  const attestText = [
    "Je soussigné(e), " + userName + ", collaborateur(trice) d'Eiffel Investment Group,",
    "atteste avoir consulté la Restricted List en date du " + dateStr,
    "à " + timeStr + ".",
    "",
    "Cette consultation a été effectuée dans le cadre de mes obligations de",
    "conformité et de prévention des opérations d'initiés.",
    "",
    "Je reconnais avoir pris connaissance des restrictions en vigueur",
    "sur les titres financiers listés."
  ];

  attestText.forEach(line => {
    doc.text(line, 30, y + 5);
    y += 7;
  });

  y += 20;

  // --- Signature ---
  doc.setDrawColor(13, 27, 42);
  doc.setLineWidth(0.3);
  doc.line(20, y + 20, 90, y + 20);
  doc.setFontSize(9);
  doc.setTextColor(100, 100, 100);
  doc.text("Signature du collaborateur", 20, y + 26);

  // --- Pied de page ---
  doc.setFillColor(13, 27, 42);
  doc.rect(0, 282, 210, 15, "F");
  doc.setFontSize(8);
  doc.setTextColor(150, 150, 150);
  doc.text("Document généré automatiquement par le Compliance Portal — Eiffel Investment Group", 105, 291, { align: "center" });
  doc.text(`Généré le ${now.toLocaleDateString("fr-FR")} à ${timeStr}`, 105, 296, { align: "center" });

  // --- Téléchargement ---
  const filename = `Attestation_RestrictedList_${now.toISOString().slice(0, 10)}.pdf`;
  doc.save(filename);
}

// -----------------------------------------------
// 11. GÉNÉRATION PDF — DÉCLARATION DE TRANSACTION
// -----------------------------------------------
function generateDeclarationPDF(data) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  const now = new Date();
  const dateDeclaration = now.toLocaleDateString("fr-FR", {
    weekday: "long", year: "numeric", month: "long", day: "numeric"
  });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;

  // --- En-tête ---
  doc.setFillColor(13, 27, 42);
  doc.rect(0, 0, 210, 40, "F");

  doc.setTextColor(201, 168, 76);
  doc.setFontSize(20);
  doc.setFont("helvetica", "bold");
  doc.text("EIFFEL INVESTMENT GROUP", 20, 18);

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.text("Compliance Portal", 20, 28);

  // --- Titre ---
  doc.setTextColor(13, 27, 42);
  doc.setFontSize(16);
  doc.setFont("helvetica", "bold");
  doc.text("ATTESTATION DE DÉCLARATION", 105, 58, { align: "center" });

  doc.setFontSize(12);
  doc.setFont("helvetica", "normal");
  doc.text("Transaction sur titre financier", 105, 68, { align: "center" });

  doc.setDrawColor(201, 168, 76);
  doc.setLineWidth(0.8);
  doc.line(20, 74, 190, 74);

  // --- Infos collaborateur ---
  let y = 88;
  doc.setFontSize(12);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(13, 27, 42);
  doc.text("INFORMATIONS DU DÉCLARANT", 20, y);
  y += 8;

  doc.setFontSize(10.5);
  const addRow = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.setTextColor(80, 80, 80);
    doc.text(label, 20, y);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(30, 30, 30);
    doc.text(String(value || "—"), 80, y);
    y += 9;
  };

  addRow("Nom :", userName);
  addRow("Email :", currentUser.mail || "—");
  addRow("Date de déclaration :", dateDeclaration);
  addRow("Heure :", timeStr);

  y += 4;
  doc.setDrawColor(200, 200, 200);
  doc.line(20, y, 190, y);
  y += 10;

  // --- Détails de la transaction ---
  doc.setFontSize(12);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(13, 27, 42);
  doc.text("DÉTAILS DE LA TRANSACTION", 20, y);
  y += 8;

  doc.setFontSize(10.5);
  addRow("Titre financier :", data.titre);
  addRow("Type d'opération :", data.type);
  addRow("Quantité :", data.quantite);
  addRow("Prix unitaire :", data.prix ? `${data.prix} €` : "Non renseigné");
  addRow("Date de transaction :", data.date ? new Date(data.date).toLocaleDateString("fr-FR") : "—");
  addRow("Compte utilisé :", data.compte);
  if (data.commentaire) {
    addRow("Commentaire :", data.commentaire);
  }

  y += 4;
  doc.line(20, y, 190, y);
  y += 12;

  // --- Texte d'attestation ---
  doc.setFillColor(248, 249, 250);
  doc.rect(20, y - 5, 170, 40, "F");
  doc.setDrawColor(220, 220, 220);
  doc.rect(20, y - 5, 170, 40);

  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(50, 50, 50);

  const lines = [
    `Je soussigné(e), ${userName}, atteste avoir déclaré la transaction`,
    `ci-dessus dans le cadre de mes obligations de conformité.`,
    `Cette déclaration a été enregistrée le ${dateDeclaration} à ${timeStr}.`,
    `Je certifie l'exactitude des informations fournies.`
  ];

  lines.forEach(line => {
    doc.text(line, 28, y + 5);
    y += 8;
  });

  y += 16;

  // --- Signature ---
  doc.setDrawColor(13, 27, 42);
  doc.line(20, y + 20, 90, y + 20);
  doc.setFontSize(9);
  doc.setTextColor(100, 100, 100);
  doc.text("Signature du collaborateur", 20, y + 26);

  // --- Pied de page ---
  doc.setFillColor(13, 27, 42);
  doc.rect(0, 282, 210, 15, "F");
  doc.setFontSize(8);
  doc.setTextColor(150, 150, 150);
  doc.text("Document généré automatiquement par le Compliance Portal — Eiffel Investment Group", 105, 291, { align: "center" });

  const filename = `Declaration_${data.titre.replace(/\s/g, "_")}_${now.toISOString().slice(0, 10)}.pdf`;
  doc.save(filename);
}

// -----------------------------------------------
// 12. FONCTION UTILITAIRE — APPEL MICROSOFT GRAPH API
// -----------------------------------------------
async function callGraphAPI(endpoint, method = "GET", body = null) {
  // Récupère un token d'accès (silencieusement si possible)
  let tokenResponse;
  try {
    tokenResponse = await msalInstance.acquireTokenSilent({
      ...graphScopes,
      account: msalInstance.getActiveAccount()
    });
  } catch {
    // Si le token silencieux échoue, on redemande en popup
    tokenResponse = await msalInstance.acquireTokenPopup(graphScopes);
  }

  // Prépare la requête
  const options = {
    method,
    headers: {
      "Authorization": `Bearer ${tokenResponse.accessToken}`,
      "Content-Type": "application/json"
    }
  };

  if (body) options.body = JSON.stringify(body);

  // Envoie la requête à l'API Graph
  const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, options);

  if (!response.ok) {
    const err = await response.json().catch(() => ({}));
    throw new Error(err.error?.message || `Erreur API: ${response.status}`);
  }

  // Si la réponse est vide (ex: après un POST), on retourne null
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

// -----------------------------------------------
// 13. FONCTIONS UTILITAIRES — AFFICHAGE
// -----------------------------------------------
function showScreen(name) {
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
  document.getElementById(`screen-${name}`).classList.add("active");
}

function showElement(id) {
  const el = document.getElementById(id);
  if (el) { el.classList.remove("hidden"); }
}

function hideElement(id) {
  const el = document.getElementById(id);
  if (el) { el.classList.add("hidden"); }
}

function showFormMessage(text, type) {
  const el = document.getElementById("declaration-msg");
  el.textContent = text;
  el.className = `form-message ${type}`;
  showElement("declaration-msg");
}

// -----------------------------------------------
// 14. LANCEMENT
// -----------------------------------------------
init();
