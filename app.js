// =============================================
//  COMPLIANCE APP — EIFFEL INVESTMENT GROUP
// =============================================

const CONFIG = {
  clientId: "56c83db2-7fc6-4826-8073-82077e91ae27",
  tenantId: "eb5288b6-5ccb-4f5d-88fb-346db1317a4f",
  redirectUri: window.location.origin,
  sharepoint: {
    siteHostname: "eiffelinvestmentgroup.sharepoint.com",
    sitePath: "/sites/CPL",
    listNDA: "NDA List",
    listInfoPriv: "Information privil\u00e9gi\u00e9e"
  }
};

const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri
  },
  cache: { cacheLocation: "sessionStorage" }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const graphScopes = { scopes: ["User.Read", "Sites.Read.All"] };

let currentUser = null;
let siteId = null;
let allRestrictedItems = [];

async function init() {
  try {
    const result = await msalInstance.handleRedirectPromise();
    if (result) {
      msalInstance.setActiveAccount(result.account);
      await onLoggedIn();
      return;
    }
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance.setActiveAccount(accounts[0]);
      await onLoggedIn();
    } else {
      showScreen("login");
    }
  } catch (err) {
    console.error("Erreur init:", err);
    showScreen("login");
  }
}

document.getElementById("btn-login").addEventListener("click", async () => {
  try {
    await msalInstance.loginRedirect(graphScopes);
  } catch (err) {
    console.error("Erreur connexion:", err);
  }
});

document.getElementById("btn-logout").addEventListener("click", () => {
  msalInstance.logoutRedirect();
});

async function onLoggedIn() {
  currentUser = await callGraphAPI("/me");
  const fullName = currentUser.displayName || currentUser.mail;
  document.getElementById("user-name").textContent = fullName;
  const initials = fullName.split(" ").map(w => w[0]).join("").substring(0, 2).toUpperCase();
  document.getElementById("user-avatar").textContent = initials;

  try {
    const site = await callGraphAPI(`/sites/${CONFIG.sharepoint.siteHostname}:${CONFIG.sharepoint.sitePath}`);
    siteId = site.id;
  } catch (err) {
    console.error("Erreur site SharePoint:", err);
    showScreen("app");
    return;
  }

  showScreen("app");
  loadRestrictedList();
  loadHistory();
}

document.querySelectorAll(".nav-tab").forEach(tab => {
  tab.addEventListener("click", () => {
    const target = tab.dataset.tab;
    document.querySelectorAll(".nav-tab").forEach(t => t.classList.remove("active"));
    tab.classList.add("active");
    document.querySelectorAll(".tab-content").forEach(s => s.classList.remove("active"));
    document.getElementById(`tab-${target}`).classList.add("active");
  });
});

async function loadRestrictedList() {
  showElement("restricted-loading");
  hideElement("restricted-table-container");
  hideElement("restricted-error");

  try {
    const [ndaData, infoPrivData] = await Promise.all([
      callGraphAPI(`/sites/${siteId}/lists/${encodeURIComponent(CONFIG.sharepoint.listNDA)}/items?expand=fields&$top=500`),
      callGraphAPI(`/sites/${siteId}/lists/${encodeURIComponent(CONFIG.sharepoint.listInfoPriv)}/items?expand=fields&$top=500`)
    ]);

    const ndaItems = (ndaData.value || []).map(item => {
      const f = item.fields || {};
      return {
        nom: f.Title || f.Soci_x00e9_t_x00e9_ || f.Societe || "—",
        isin: f.ISIN || f.CodeISIN || "—",
        dateDebut: f.Date_x0020_de_x0020_d_x00e9_but || f.Datedebut || f.DateDebut || "—",
        dateFin: f.Date_x0020_de_x0020_fin || f.Datefin || f.DateFin || "—",
        type: "NDA",
        source: "NDA List"
      };
    });

    const infoPrivItems = (infoPrivData.value || []).map(item => {
      const f = item.fields || {};
      return {
        nom: f.Title || f.Titre || "—",
        isin: f.CodeISIN || f.Code_x0020_ISIN || f.ISIN || "—",
        dateDebut: f.Datedebut || f.DateDebut || f.Date_x0020_de_x0020_d_x00e9_but || "—",
        dateFin: f.Datefin || f.DateFin || f.Date_x0020_de_x0020_fin || "—",
        type: "Information privil\u00e9gi\u00e9e",
        source: "Info Priv"
      };
    });

    allRestrictedItems = [...ndaItems, ...infoPrivItems];
    hideElement("restricted-loading");
    renderRestrictedTable(allRestrictedItems);
    showElement("restricted-table-container");

  } catch (err) {
    console.error("Erreur chargement:", err);
    hideElement("restricted-loading");
    showElement("restricted-error");
    document.getElementById("restricted-error-msg").textContent = "Impossible de charger la liste. V\u00e9rifiez vos permissions SharePoint.";
  }
}

function renderRestrictedTable(items) {
  document.getElementById("restricted-count").textContent = `${items.length} titre${items.length > 1 ? "s" : ""}`;
  const tbody = document.getElementById("restricted-tbody");
  tbody.innerHTML = "";

  if (items.length === 0) {
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center;color:#9aaaba;padding:32px">Aucun r\u00e9sultat</td></tr>`;
    return;
  }

  items.forEach(item => {
    const badgeClass = item.source === "NDA List" ? "badge-nda" : "badge-restricted";
    tbody.innerHTML += `
      <tr>
        <td><strong>${item.nom}</strong></td>
        <td style="font-family:monospace;font-size:13px">${item.isin}</td>
        <td><span class="badge ${badgeClass}">${item.type}</span></td>
        <td>${formatDate(item.dateDebut)}</td>
        <td>${formatDate(item.dateFin)}</td>
      </tr>`;
  });
}

document.getElementById("search-restricted").addEventListener("input", function () {
  const query = this.value.toLowerCase().trim();
  const filtered = query
    ? allRestrictedItems.filter(item =>
        item.nom.toLowerCase().includes(query) ||
        item.isin.toLowerCase().includes(query))
    : allRestrictedItems;
  renderRestrictedTable(filtered);
});

document.getElementById("search-history").addEventListener("input", function () {
  renderHistory(this.value.toLowerCase().trim());
});

document.getElementById("btn-generate-restricted-pdf").addEventListener("click", () => {
  if (!currentUser) return;
  generateRestrictedPDF();
});

function generateRestrictedPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const now = new Date();
  const dateStr = now.toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;

  doc.setFillColor(0, 20, 59);
  doc.rect(0, 0, 210, 40, "F");
  doc.setTextColor(209, 143, 65);
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.text("EIFFEL INVESTMENT GROUP", 20, 16);
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.text("Compliance Portal", 20, 26);

  doc.setTextColor(0, 20, 59);
  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.text("ATTESTATION DE CONSULTATION", 105, 55, { align: "center" });
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 80, 80);
  doc.text("Restricted List \u2014 Titres sous restriction", 105, 63, { align: "center" });

  doc.setDrawColor(209, 143, 65);
  doc.setLineWidth(0.8);
  doc.line(20, 68, 190, 68);

  let y = 82;
  const addLine = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 20, 59);
    doc.setFontSize(10.5);
    doc.text(label, 20, y);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(50, 50, 50);
    doc.text(value, 75, y);
    y += 9;
  };

  addLine("Collaborateur :", userName);
  addLine("Email :", currentUser.mail || "—");
  addLine("Date de consultation :", dateStr);
  addLine("Heure :", timeStr);
  addLine("Titres consult\u00e9s :", `${allRestrictedItems.length} titre(s)`);

  y += 6;
  doc.setFillColor(248, 249, 250);
  doc.rect(20, y - 4, 170, 50, "F");
  doc.setDrawColor(220, 220, 220);
  doc.rect(20, y - 4, 170, 50);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(50, 50, 50);
  [`Je soussign\u00e9(e), ${userName}, collaborateur(trice) d'Eiffel Investment Group,`,
   `atteste avoir consult\u00e9 la Restricted List en date du ${dateStr}`,
   `\u00e0 ${timeStr}.`,
   "",
   "Cette consultation a \u00e9t\u00e9 effectu\u00e9e dans le cadre de mes obligations de",
   "conformit\u00e9 et de pr\u00e9vention des op\u00e9rations d'initi\u00e9s.",
   "",
   "Je reconnais avoir pris connaissance des restrictions en vigueur."
  ].forEach(line => { doc.text(line, 28, y + 6); y += 7; });

  y += 14;
  doc.setDrawColor(0, 20, 59);
  doc.line(20, y + 20, 90, y + 20);
  doc.setFontSize(9);
  doc.setTextColor(100, 100, 100);
  doc.text("Signature du collaborateur", 20, y + 27);

  doc.setFillColor(0, 20, 59);
  doc.rect(0, 282, 210, 15, "F");
  doc.setFontSize(8);
  doc.setTextColor(150, 150, 150);
  doc.text("Document g\u00e9n\u00e9r\u00e9 automatiquement \u2014 Eiffel Investment Group Compliance Portal", 105, 291, { align: "center" });

  doc.save(`Attestation_RestrictedList_${now.toISOString().slice(0, 10)}.pdf`);
}

document.getElementById("btn-submit-declaration").addEventListener("click", () => {
  const titre = document.getElementById("f-titre").value.trim();
  const type = document.getElementById("f-type").value;
  const quantite = document.getElementById("f-quantite").value;
  const prix = document.getElementById("f-prix").value;
  const date = document.getElementById("f-date").value;
  const compte = document.getElementById("f-compte").value.trim();
  const commentaire = document.getElementById("f-commentaire").value.trim();

  if (!titre || !type || !quantite || !date || !compte) {
    showFormMessage("Veuillez remplir tous les champs obligatoires (*)", "error");
    return;
  }

  const data = { titre, type, quantite, prix, date, compte, commentaire };
  generateDeclarationPDF(data);
  saveToHistory(data);
  showFormMessage("\u2713 Attestation PDF g\u00e9n\u00e9r\u00e9e ! N'oubliez pas de l'envoyer au service Compliance.", "success");

  setTimeout(() => {
    ["f-titre","f-type","f-quantite","f-prix","f-date","f-compte","f-commentaire"].forEach(id => {
      document.getElementById(id).value = "";
    });
    hideElement("declaration-msg");
    loadHistory();
  }, 4000);
});

function generateDeclarationPDF(data) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const now = new Date();
  const dateDeclaration = now.toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;

  doc.setFillColor(0, 20, 59);
  doc.rect(0, 0, 210, 40, "F");
  doc.setTextColor(209, 143, 65);
  doc.setFontSize(18);
  doc.setFont("helvetica", "bold");
  doc.text("EIFFEL INVESTMENT GROUP", 20, 16);
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.text("Compliance Portal", 20, 26);

  doc.setTextColor(0, 20, 59);
  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.text("ATTESTATION DE D\u00c9CLARATION", 105, 55, { align: "center" });
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 80, 80);
  doc.text("Transaction sur titre financier", 105, 63, { align: "center" });

  doc.setDrawColor(209, 143, 65);
  doc.setLineWidth(0.8);
  doc.line(20, 68, 190, 68);

  let y = 80;
  doc.setFontSize(11);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(0, 20, 59);
  doc.text("D\u00c9CLARANT", 20, y);
  y += 8;

  const addRow = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.setTextColor(80, 80, 80);
    doc.text(label, 20, y);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(30, 30, 30);
    doc.text(String(value || "\u2014"), 75, y);
    y += 8;
  };

  addRow("Nom :", userName);
  addRow("Email :", currentUser.mail || "\u2014");
  addRow("Date de d\u00e9claration :", dateDeclaration);
  addRow("Heure :", timeStr);

  y += 4;
  doc.setDrawColor(200, 200, 200);
  doc.line(20, y, 190, y);
  y += 8;

  doc.setFontSize(11);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(0, 20, 59);
  doc.text("D\u00c9TAILS DE LA TRANSACTION", 20, y);
  y += 8;

  addRow("Titre financier :", data.titre);
  addRow("Type d'op\u00e9ration :", data.type);
  addRow("Quantit\u00e9 :", data.quantite);
  addRow("Prix unitaire :", data.prix ? `${data.prix} \u20ac` : "Non renseign\u00e9");
  addRow("Date de transaction :", data.date ? new Date(data.date).toLocaleDateString("fr-FR") : "\u2014");
  addRow("Compte utilis\u00e9 :", data.compte);
  if (data.commentaire) addRow("Commentaire :", data.commentaire);

  y += 4;
  doc.line(20, y, 190, y);
  y += 10;

  doc.setFillColor(248, 249, 250);
  doc.rect(20, y - 4, 170, 36, "F");
  doc.setDrawColor(220, 220, 220);
  doc.rect(20, y - 4, 170, 36);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(50, 50, 50);
  [`Je soussign\u00e9(e), ${userName}, atteste avoir d\u00e9clar\u00e9 la transaction`,
   `ci-dessus dans le cadre de mes obligations de conformit\u00e9.`,
   `Cette d\u00e9claration a \u00e9t\u00e9 g\u00e9n\u00e9r\u00e9e le ${dateDeclaration} \u00e0 ${timeStr}.`,
   `Je certifie l'exactitude des informations fournies.`
  ].forEach(line => { doc.text(line, 28, y + 5); y += 8; });

  y += 10;

  doc.setFillColor(255, 243, 220);
  doc.rect(20, y, 170, 20, "F");
  doc.setDrawColor(209, 143, 65);
  doc.setLineWidth(0.8);
  doc.rect(20, y, 170, 20);
  doc.setFontSize(10);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(133, 79, 11);
  doc.text("RAPPEL IMPORTANT", 28, y + 7);
  doc.setFont("helvetica", "normal");
  doc.text("Cette attestation doit \u00eatre transmise au service Compliance dans les 48 heures.", 28, y + 15);

  y += 30;
  doc.setDrawColor(0, 20, 59);
  doc.setLineWidth(0.3);
  doc.line(20, y + 20, 90, y + 20);
  doc.setFontSize(9);
  doc.setTextColor(100, 100, 100);
  doc.text("Signature du collaborateur", 20, y + 27);

  doc.setFillColor(0, 20, 59);
  doc.rect(0, 282, 210, 15, "F");
  doc.setFontSize(8);
  doc.setTextColor(150, 150, 150);
  doc.text("Document g\u00e9n\u00e9r\u00e9 automatiquement \u2014 Eiffel Investment Group Compliance Portal", 105, 291, { align: "center" });

  doc.save(`Declaration_${data.titre.replace(/\s/g,"_")}_${now.toISOString().slice(0,10)}.pdf`);
}

function getHistoryKey() {
  return `eig_declarations_${currentUser.mail}`;
}

function saveToHistory(data) {
  const key = getHistoryKey();
  const existing = JSON.parse(localStorage.getItem(key) || "[]");
  existing.unshift({ ...data, dateDeclaration: new Date().toISOString(), id: Date.now() });
  localStorage.setItem(key, JSON.stringify(existing));
}

function loadHistory() {
  if (!currentUser) return;
  renderHistory("");
}

function renderHistory(query) {
  const key = getHistoryKey();
  let items = JSON.parse(localStorage.getItem(key) || "[]");
  hideElement("history-loading");

  if (query) items = items.filter(i => i.titre.toLowerCase().includes(query));

  if (items.length === 0) {
    showElement("history-empty");
    hideElement("history-table-container");
    return;
  }

  hideElement("history-empty");
  showElement("history-table-container");

  const tbody = document.getElementById("history-tbody");
  tbody.innerHTML = "";
  items.forEach(item => {
    const dateDecl = new Date(item.dateDeclaration).toLocaleDateString("fr-FR");
    const dateTx = item.date ? new Date(item.date).toLocaleDateString("fr-FR") : "\u2014";
    tbody.innerHTML += `
      <tr>
        <td><strong>${item.titre}</strong></td>
        <td>${item.type}</td>
        <td>${item.quantite}</td>
        <td>${item.prix ? item.prix + " \u20ac" : "\u2014"}</td>
        <td>${dateTx}</td>
        <td>${dateDecl}</td>
        <td><button class="btn-pdf" onclick='regenPDF(${JSON.stringify(item).replace(/'/g,"&#39;")})'>
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          PDF
        </button></td>
      </tr>`;
  });
}

function regenPDF(item) { generateDeclarationPDF(item); }

async function callGraphAPI(endpoint, method = "GET", body = null) {
  let tokenResponse;
  try {
    tokenResponse = await msalInstance.acquireTokenSilent({ ...graphScopes, account: msalInstance.getActiveAccount() });
  } catch {
    await msalInstance.acquireTokenRedirect(graphScopes);
    return;
  }
  const options = {
    method,
    headers: { "Authorization": `Bearer ${tokenResponse.accessToken}`, "Content-Type": "application/json" }
  };
  if (body) options.body = JSON.stringify(body);
  const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, options);
  if (!response.ok) { const err = await response.json().catch(() => ({})); throw new Error(err.error?.message || `Erreur API: ${response.status}`); }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

function showScreen(name) {
  document.querySelectorAll(".screen").forEach(s => s.classList.remove("active"));
  document.getElementById(`screen-${name}`).classList.add("active");
}
function showElement(id) { const el = document.getElementById(id); if (el) el.classList.remove("hidden"); }
function hideElement(id) { const el = document.getElementById(id); if (el) el.classList.add("hidden"); }
function showFormMessage(text, type) {
  const el = document.getElementById("declaration-msg");
  el.textContent = text;
  el.className = `form-message ${type}`;
  showElement("declaration-msg");
}
function formatDate(dateStr) {
  if (!dateStr || dateStr === "\u2014") return "\u2014";
  try { return new Date(dateStr).toLocaleDateString("fr-FR"); } catch { return dateStr; }
}

init();
