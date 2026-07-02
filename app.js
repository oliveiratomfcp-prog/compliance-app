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
  },
  sharepointHistory: {
    sitePath: "/sites/CPLDashboard",
    listHistory: "Historique Compliance",
    listGifts: "Registre cadeaux"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const graphScopes = { scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"] };

let currentUser = null;
let siteId = null;
let siteIdHistory = null;
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
    console.error("Erreur site SharePoint CPL:", err);
    showScreen("app");
    return;
  }

  try {
    const site2 = await callGraphAPI(`/sites/${CONFIG.sharepoint.siteHostname}:${CONFIG.sharepointHistory.sitePath}`);
    siteIdHistory = site2.id;
  } catch (err) {
    console.error("Erreur site SharePoint CPLDashboard:", err);
  }

  showScreen("app");
  loadRestrictedList();
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

// Charge tous les items d'une liste avec pagination
async function getAllListItems(siteId, listName, fieldsExpand) {
  let items = [];
  const expandParam = fieldsExpand
    ? `fields($expand=${fieldsExpand})`
    : "fields";
  let url = `/sites/${siteId}/lists/${encodeURIComponent(listName)}/items?$expand=${expandParam}&$top=500`;
  while (url) {
    const data = await callGraphAPI(url);
    items = items.concat(data.value || []);
    url = data["@odata.nextLink"]
      ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
      : null;
  }
  return items;
}

// Retourne true si la date de fin est dans le passé
function isExpired(dateFin) {
  if (!dateFin || dateFin === "—") return false;
  return new Date(dateFin) < new Date();
}

async function loadRestrictedList() {
  showElement("restricted-loading");
  hideElement("restricted-table-container");
  hideElement("restricted-error");

  try {
    const [ndaRaw, infoPrivRaw] = await Promise.all([
      getAllListItems(siteId, CONFIG.sharepoint.listNDA),
      getAllListItems(siteId, CONFIG.sharepoint.listInfoPriv)
    ]);

    const ndaItems = (ndaRaw || []).map(item => {
      const f = item.fields || {};
      const sigName = f['SignataireNom'] || '—';
      return {
        nom: f.Title || "—",
        isin: f.ISIN || f.CodeISIN || "—",
        dateDebut: f.Dateded_x00e9_but || "—",
        dateFin: f.Datedefin || "—",
        equipe: "—",
        signataire: sigName,
        type: "NDA",
        source: "NDA List"
      };
    }).filter(item => !isExpired(item.dateFin));

    const infoPrivItems = (infoPrivRaw || []).map(item => {
      const f = item.fields || {};
      return {
        nom: f.Title || "—",
        isin: f.CodeISIN || "—",
        dateDebut: f.Dateded_x00e9_but || "—",
        dateFin: f.Datedefin || "—",
        equipe: f.Equipe || "—",
        signataire: "—",
        type: "Information privil\u00e9gi\u00e9e",
        source: "Info Priv"
      };
    }).filter(item => !isExpired(item.dateFin));

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
    tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:#9aaaba;padding:32px">Aucun r\u00e9sultat</td></tr>`;
    return;
  }

  const fragment = document.createDocumentFragment();
  items.forEach(item => {
    const badgeClass = item.source === "NDA List" ? "badge-nda" : "badge-restricted";
    const isFutureEnd = item.dateFin && item.dateFin !== "\u2014" && new Date(item.dateFin) > new Date();
    const dateFinCell = isFutureEnd
      ? `<span class="pill-live"><span class="live-dot"></span>En cours</span>`
      : formatDate(item.dateFin);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${item.nom}</strong></td>
      <td style="font-family:monospace;font-size:13px">${item.isin}</td>
      <td><span class="badge ${badgeClass}">${item.type}</span></td>
      <td>${item.equipe && item.equipe !== "\u2014" ? `<span class="badge badge-active">${item.equipe}</span>` : "\u2014"}</td>
      <td>${item.signataire !== "\u2014" ? item.signataire : "\u2014"}</td>
      <td>${formatDate(item.dateDebut)}</td>
      <td>${dateFinCell}</td>`;
    fragment.appendChild(tr);
  });
  tbody.appendChild(fragment);
}

let currentSort = { col: null, dir: 1 };

function sortTable(col) {
  if (currentSort.col === col) {
    currentSort.dir *= -1;
  } else {
    currentSort.col = col;
    currentSort.dir = 1;
  }

  ["nom", "isin", "type", "equipe", "signataire", "dateDebut", "dateFin"].forEach(c => {
    const el = document.getElementById(`sort-${c}`);
    const th = el ? el.parentElement : null;
    if (el) el.textContent = "\u2195";
    if (th) th.classList.remove("sort-asc", "sort-desc");
  });
  const activeEl = document.getElementById(`sort-${col}`);
  if (activeEl) {
    activeEl.textContent = currentSort.dir === 1 ? "\u2191" : "\u2193";
    activeEl.parentElement.classList.add(currentSort.dir === 1 ? "sort-asc" : "sort-desc");
  }

  const query = document.getElementById("search-restricted").value.toLowerCase().trim();
  let items = query
    ? allRestrictedItems.filter(item =>
        item.nom.toLowerCase().includes(query) ||
        item.isin.trim().toLowerCase().includes(query))
    : [...allRestrictedItems];

  items.sort((a, b) => {
    let valA = a[col] || "";
    let valB = b[col] || "";
    if (col === "dateDebut" || col === "dateFin") {
      valA = valA ? new Date(valA).getTime() : 0;
      valB = valB ? new Date(valB).getTime() : 0;
      return (valA - valB) * currentSort.dir;
    }
    return valA.localeCompare(valB, "fr") * currentSort.dir;
  });

  renderRestrictedTable(items);
}

let searchTimeout = null;

document.getElementById("search-restricted").addEventListener("input", function () {
  clearTimeout(searchTimeout);
  const query = this.value.toLowerCase().trim();
  searchTimeout = setTimeout(() => {
    currentSort = { col: null, dir: 1 };
    ["nom", "isin", "type", "equipe", "signataire", "dateDebut", "dateFin"].forEach(c => {
      const el = document.getElementById(`sort-${c}`);
      if (el) { el.textContent = "\u2195"; el.parentElement.classList.remove("sort-asc", "sort-desc"); }
    });
    const filtered = query
      ? allRestrictedItems.filter(item =>
          item.nom.toLowerCase().includes(query) ||
          item.isin.trim().toLowerCase().includes(query))
      : allRestrictedItems;
    renderRestrictedTable(filtered);
  }, 200);
});

document.getElementById("search-history").addEventListener("input", function () {
  renderHistory(this.value.toLowerCase().trim());
});

// -----------------------------------------------
// EXPORT EXCEL
// -----------------------------------------------
document.getElementById("btn-export-excel").addEventListener("click", () => {
  const query = document.getElementById("search-restricted").value.toLowerCase().trim();
  const items = query
    ? allRestrictedItems.filter(i => i.nom.toLowerCase().includes(query) || i.isin.trim().toLowerCase().includes(query))
    : allRestrictedItems;

  const bom = "\uFEFF";
  const headers = ["Soci\u00e9t\u00e9 / Titre", "Code ISIN", "Type de restriction", "\u00c9quipe", "Date de d\u00e9but", "Date de fin"];
  const rows = items.map(item => [
    `"${item.nom}"`, `"${item.isin}"`, `"${item.type}"`,
    `"${item.equipe || ""}"`, `"${formatDate(item.dateDebut)}"`, `"${formatDate(item.dateFin)}"`
  ].join(";"));

  const csv = bom + headers.join(";") + "\n" + rows.join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `RestrictedList_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
});

// -----------------------------------------------
// VÉRIFICATION EN MASSE
// -----------------------------------------------
document.getElementById("btn-bulk-check").addEventListener("click", () => {
  const container = document.getElementById("bulk-check-container");
  container.classList.toggle("hidden");
  if (!container.classList.contains("hidden")) {
    document.getElementById("bulk-input").focus();
  }
});

document.getElementById("btn-bulk-run").addEventListener("click", () => {
  const raw = document.getElementById("bulk-input").value.trim();
  if (!raw) return;

  const entries = raw.split(/[\n,;]+/).map(e => e.trim()).filter(e => e.length > 0);
  if (entries.length === 0) return;

  const resultsEl = document.getElementById("bulk-results");
  resultsEl.innerHTML = "";
  resultsEl.classList.remove("hidden");

  let restricted = 0;
  let clean = 0;
  const resultData = [];

  entries.forEach(entry => {
    const entryClean = entry.trim().toLowerCase();
    const matches = allRestrictedItems.filter(item =>
      item.nom.toLowerCase().includes(entryClean) ||
      item.isin.trim().toLowerCase().includes(entryClean) ||
      entryClean.includes(item.isin.trim().toLowerCase())
    );
    const isRestricted = matches.length > 0;
    if (isRestricted) restricted++; else clean++;
    resultData.push({ entry, isRestricted, matches });

    const div = document.createElement("div");
    div.className = `bulk-result-item ${isRestricted ? "restricted" : "clean"}`;
    div.innerHTML = `
      <span class="bulk-result-icon">${isRestricted ? "\uD83D\uDEAB" : "\u2705"}</span>
      <div style="flex:1">
        <strong>${entry}</strong>
        ${isRestricted
          ? `<span style="margin-left:8px;font-size:12px;color:#c53030">Restreint \u2014 ${matches.map(m => m.type).join(", ")}</span>`
          : `<span style="margin-left:8px;font-size:12px;color:var(--success)">Non restreint</span>`}
      </div>`;
    resultsEl.appendChild(div);
  });

  const summary = document.createElement("div");
  summary.style.cssText = "display:flex;gap:10px;margin-top:12px;padding-top:12px;border-top:1px solid var(--border);justify-content:space-between;align-items:center";
  summary.innerHTML = `
    <span style="font-size:13px;color:var(--text-secondary)">
      <strong>${entries.length}</strong> titre(s) \u2014
      <span style="color:#c53030"><strong>${restricted}</strong> restreint(s)</span>,
      <span style="color:var(--success)"><strong>${clean}</strong> autoris\u00e9(s)</span>
    </span>
    <button onclick="generateBulkPDF(${JSON.stringify(resultData).replace(/'/g,'&#39;').replace(/"/g,'&quot;')})" class="btn-primary" style="font-size:12px;padding:8px 14px">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
      Rapport PDF
    </button>`;
  resultsEl.appendChild(summary);
});

function generateBulkPDF(resultData) {
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
  doc.setFontSize(14);
  doc.setFont("helvetica", "bold");
  doc.text("RAPPORT DE V\u00c9RIFICATION EN MASSE", 105, 55, { align: "center" });
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 80, 80);
  doc.text("Restricted List \u2014 V\u00e9rification multiple", 105, 63, { align: "center" });

  doc.setDrawColor(209, 143, 65);
  doc.setLineWidth(0.8);
  doc.line(20, 68, 190, 68);

  let y = 80;
  doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59); doc.setFontSize(10.5);
  doc.text("Collaborateur :", 20, y); doc.setFont("helvetica", "normal"); doc.setTextColor(50,50,50); doc.text(userName, 80, y); y += 9;
  doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59);
  doc.text("Date :", 20, y); doc.setFont("helvetica", "normal"); doc.setTextColor(50,50,50); doc.text(dateStr, 80, y); y += 9;
  doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59);
  doc.text("Heure :", 20, y); doc.setFont("helvetica", "normal"); doc.setTextColor(50,50,50); doc.text(timeStr, 80, y); y += 9;
  const restr = resultData.filter(r => r.isRestricted).length;
  doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59);
  doc.text("R\u00e9sum\u00e9 :", 20, y); doc.setFont("helvetica", "normal"); doc.setTextColor(50,50,50);
  doc.text(`${resultData.length} titre(s) \u2014 ${restr} restreint(s), ${resultData.length - restr} autoris\u00e9(s)`, 80, y); y += 14;

  doc.setDrawColor(200,200,200); doc.line(20, y, 190, y); y += 10;
  doc.setFontSize(11); doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59);
  doc.text("D\u00c9TAIL", 20, y); y += 10;

  resultData.forEach(r => {
    if (y > 265) { doc.addPage(); y = 20; }
    if (r.isRestricted) {
      doc.setFillColor(253, 232, 232); doc.rect(20, y-4, 170, 16, "F");
      doc.setDrawColor(245,198,198); doc.rect(20, y-4, 170, 16);
      doc.setFont("helvetica", "bold"); doc.setTextColor(197,48,48); doc.setFontSize(10);
      doc.text(`\u26a0 ${r.entry}`, 25, y+4);
      doc.setFont("helvetica", "normal"); doc.setFontSize(9);
      doc.text(`Restreint \u2014 ${r.matches.map(m => m.type).join(", ")}`, 25, y+11);
      y += 20;
    } else {
      doc.setFillColor(232,244,244); doc.rect(20, y-4, 170, 12, "F");
      doc.setFont("helvetica", "normal"); doc.setTextColor(0,102,96); doc.setFontSize(10);
      doc.text(`\u2713 ${r.entry} \u2014 Non restreint`, 25, y+4);
      y += 16;
    }
  });

  doc.setFillColor(0,20,59); doc.rect(0,282,210,15,"F");
  doc.setFontSize(8); doc.setTextColor(150,150,150);
  doc.text("Document g\u00e9n\u00e9r\u00e9 automatiquement \u2014 Eiffel Investment Group Compliance Portal", 105, 291, { align: "center" });

  saveConsultationToHistory(
    `V\u00e9rification masse (${resultData.length} titres)`,
    restr > 0,
    resultData.filter(r => r.isRestricted).flatMap(r => r.matches)
  );

  doc.save(`Rapport_Verification_${now.toISOString().slice(0,10)}.pdf`);
}

document.getElementById("btn-generate-restricted-pdf").addEventListener("click", () => {
  if (!currentUser) return;
  const query = document.getElementById("search-restricted").value.trim();
  generateRestrictedPDF(query);
});

function generateRestrictedPDF(query) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const now = new Date();
  const dateStr = now.toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;

  // Cherche si le titre est dans la restricted list
  const found = query
    ? allRestrictedItems.filter(item =>
        item.nom.toLowerCase().includes(query.toLowerCase()) ||
        item.isin.trim().toLowerCase().includes(query.toLowerCase()))
    : [];

  const isFound = found.length > 0;
  const searchLabel = query || "Consultation générale";

  // En-tête
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

  // Titre
  doc.setTextColor(0, 20, 59);
  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.text("ATTESTATION DE CONSULTATION", 105, 55, { align: "center" });
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 80, 80);
  doc.text("Restricted List \u2014 V\u00e9rification de titre", 105, 63, { align: "center" });

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
    doc.text(value, 80, y);
    y += 9;
  };

  addLine("Collaborateur :", userName);
  addLine("Email :", currentUser.mail || "\u2014");
  addLine("Date de consultation :", dateStr);
  addLine("Heure :", timeStr);
  addLine("Terme recherch\u00e9 :", searchLabel);
  addLine("R\u00e9sultat :", isFound ? `${found.length} titre(s) trouv\u00e9(s)` : "Aucun r\u00e9sultat");

  y += 4;

  // Bloc résultat coloré
  if (isFound) {
    doc.setFillColor(253, 232, 232);
    doc.rect(20, y - 4, 170, 22, "F");
    doc.setDrawColor(197, 48, 48);
    doc.setLineWidth(0.8);
    doc.rect(20, y - 4, 170, 22);
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(197, 48, 48);
    doc.text("\u26a0 TITRE SOUS RESTRICTION", 28, y + 5);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(`Le titre "${searchLabel}" figure sur la Restricted List d'Eiffel Investment Group.`, 28, y + 14);
    y += 30;

    // Détail des titres trouvés
    doc.setFontSize(10.5);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 20, 59);
    doc.text("D\u00e9tail des titres trouv\u00e9s :", 20, y);
    y += 8;

    found.forEach(item => {
      doc.setFont("helvetica", "normal");
      doc.setTextColor(50, 50, 50);
      doc.setFontSize(10);
      doc.text(`\u2022 ${item.nom}${item.isin !== "\u2014" ? ` (ISIN: ${item.isin})` : ""} \u2014 ${item.type}`, 24, y);
      doc.text(`  Date de fin : ${formatDate(item.dateFin)}`, 24, y + 6);
      y += 14;
    });
  } else {
    doc.setFillColor(232, 244, 244);
    doc.rect(20, y - 4, 170, 22, "F");
    doc.setDrawColor(0, 102, 96);
    doc.setLineWidth(0.8);
    doc.rect(20, y - 4, 170, 22);
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 102, 96);
    doc.text("\u2713 TITRE NON RESTREINT", 28, y + 5);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(`Le titre "${searchLabel}" n'appara\u00eet pas sur la Restricted List \u00e0 cette date.`, 28, y + 14);
    y += 30;
  }

  y += 6;
  doc.setFillColor(248, 249, 250);
  doc.rect(20, y - 4, 170, 30, "F");
  doc.setDrawColor(220, 220, 220);
  doc.rect(20, y - 4, 170, 30);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(50, 50, 50);
  [`Je soussign\u00e9(e), ${userName}, atteste avoir effectu\u00e9 cette v\u00e9rification`,
   `sur la Restricted List d'Eiffel Investment Group le ${dateStr} \u00e0 ${timeStr}.`,
   "Ce document constitue une preuve de consultation \u00e0 la date et l'heure indiqu\u00e9es."
  ].forEach(line => { doc.text(line, 28, y + 6); y += 8; });

  y += 14;
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

  const filename = `Attestation_${query ? query.replace(/\s/g,"_") : "RestrictedList"}_${now.toISOString().slice(0,10)}.pdf`;
  doc.save(filename);

  // Sauvegarde dans l'historique
  saveConsultationToHistory(searchLabel, isFound, found);
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

  // Vérifie si le titre est dans la restricted list
  const restricted = allRestrictedItems.filter(item =>
    item.nom.toLowerCase().includes(titre.toLowerCase()) ||
    item.isin.trim().toLowerCase().includes(titre.toLowerCase())
  );

  if (restricted.length > 0) {
    showFormMessage(
      `\u26a0 "${titre}" figure sur la Restricted List \u2014 cette transaction ne peut pas \u00eatre demand\u00e9e. Contactez le service Compliance.`,
      "error"
    );
    return;
  }

  const data = { titre, type, quantite, prix, date, compte, commentaire };

  // Génère le PDF
  generateDeclarationPDF(data);

  // Sauvegarde dans l'historique local
  saveToHistory(data);

  // Ouvre Outlook avec mail pré-rempli vers CPL
  const userName = currentUser.displayName || currentUser.mail;
  const dateFr = date ? new Date(date).toLocaleDateString("fr-FR") : "\u2014";
  const subject = encodeURIComponent(`Demande de transaction personnelle \u2014 ${titre} \u2014 ${userName}`);
  const body = encodeURIComponent(
    `Bonjour,\n\nVeuillez trouver ci-joint mon attestation de demande de transaction personnelle.\n\n` +
    `D\u00e9clarant : ${userName}\n` +
    `Titre : ${titre}\n` +
    `Type d'op\u00e9ration : ${type}\n` +
    `Quantit\u00e9 : ${quantite}\n` +
    `Prix unitaire : ${prix ? prix + " \u20ac" : "Non renseign\u00e9"}\n` +
    `Date souhait\u00e9e : ${dateFr}\n` +
    `Compte : ${compte}\n` +
    `${commentaire ? "Commentaire : " + commentaire + "\n" : ""}` +
    `\nCordialement,\n${userName}`
  );

  setTimeout(() => {
    window.location.href = `mailto:cpl@eiffel-ig.com?subject=${subject}&body=${body}`;
  }, 1000);

  showFormMessage("\u2713 PDF g\u00e9n\u00e9r\u00e9 ! Votre client mail va s'ouvrir pour envoyer l'attestation \u00e0 CPL.", "success");

  // On remet juste le message à zéro après 6 secondes mais on garde le formulaire ouvert
  setTimeout(() => {
    hideElement("declaration-msg");
  }, 6000);
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
  doc.text("DEMANDE DE TRANSACTION PERSONNELLE", 105, 55, { align: "center" });
  doc.setFontSize(11);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 80, 80);
  doc.text("Titre financier \u2014 \u00e0 transmettre \u00e0 CPL sous 24h", 105, 63, { align: "center" });

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
  addRow("Date souhait\u00e9e :", data.date ? new Date(data.date).toLocaleDateString("fr-FR") : "\u2014");
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
  doc.text("Cette demande doit \u00eatre transmise au service Compliance (CPL) dans les 24 heures.", 28, y + 15);

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

async function saveToHistory(data) {
  if (siteIdHistory) {
    try {
      await callGraphAPI(
        `/sites/${siteIdHistory}/lists/${encodeURIComponent(CONFIG.sharepointHistory.listHistory)}/items`,
        "POST",
        { fields: {
          Title: data.titre,
          D_x00e9_clarant: currentUser.displayName || currentUser.mail,
          DeclarantEmail: currentUser.mail,
          TypeEntree: "declaration",
          TypeOperation: data.type || "",
          Quantite: data.quantite ? parseFloat(data.quantite) : 0,
          Prix: data.prix ? parseFloat(data.prix) : 0,
          DateTransaction: data.date ? new Date(data.date + "T00:00:00").toISOString() : null,
          CompteUtilise: data.compte || "",
          Commentaire: data.commentaire || ""
        }}
      );
    } catch (err) {
      console.error("Erreur sauvegarde historique SharePoint:", err);
      saveToLocalStorage(data, "declaration");
    }
  } else {
    saveToLocalStorage(data, "declaration");
  }
}

async function saveConsultationToHistory(searchLabel, isFound, foundItems) {
  if (siteIdHistory) {
    try {
      await callGraphAPI(
        `/sites/${siteIdHistory}/lists/${encodeURIComponent(CONFIG.sharepointHistory.listHistory)}/items`,
        "POST",
        { fields: {
          Title: searchLabel,
          D_x00e9_clarant: currentUser.displayName || currentUser.mail,
          DeclarantEmail: currentUser.mail,
          TypeEntree: "consultation",
          Commentaire: isFound ? "Restreint - " + foundItems.map(i => i.nom).join(", ") : "Non restreint"
        }}
      );
    } catch (err) {
      console.error("Erreur sauvegarde consultation SharePoint:", err);
      saveToLocalStorage({ titre: searchLabel, isFound, foundItems }, "consultation");
    }
  } else {
    saveToLocalStorage({ titre: searchLabel, isFound, foundItems }, "consultation");
  }
}

function saveToLocalStorage(data, type) {
  const key = getHistoryKey();
  const existing = JSON.parse(localStorage.getItem(key) || "[]");
  existing.unshift({ ...data, type_entree: type, dateDeclaration: new Date().toISOString(), id: Date.now() });
  localStorage.setItem(key, JSON.stringify(existing));
}

function toggleHistory() {
  const container = document.getElementById("history-container");
  const isHidden = container.classList.contains("hidden");
  container.classList.toggle("hidden");
  if (isHidden) loadHistory();
}

let historyCache = null;

function loadHistory() {
  if (!currentUser) return;
  // Si on a déjà chargé l'historique, on réutilise le cache
  if (historyCache !== null) {
    _renderHistoryRows(historyCache);
    hideElement("history-loading");
    if (historyCache.length === 0) {
      showElement("history-empty");
      hideElement("history-table-container");
    } else {
      hideElement("history-empty");
      showElement("history-table-container");
    }
    return;
  }
  renderHistory("");
}

function refreshHistory() {
  historyCache = null;
  renderHistory("");
}

async function renderHistory(query) {
  showElement("history-loading");
  hideElement("history-table-container");
  hideElement("history-empty");

  if (siteIdHistory) {
    try {
      const email = currentUser.mail;
      const filter = `fields/DeclarantEmail eq '${email}'`;
      const data = await callGraphAPI(
        `/sites/${siteIdHistory}/lists/${encodeURIComponent(CONFIG.sharepointHistory.listHistory)}/items?expand=fields&$filter=${encodeURIComponent(filter)}&$top=100`,
        "GET", null, { "Prefer": "HonorNonIndexedQueriesWarningMayFailRandomly" }
      );

      let items = (data.value || []).map(item => {
        const f = item.fields || {};
        return {
          titre: f.Title || "\u2014",
          type_entree: f.TypeEntree || "declaration",
          type: f.TypeOperation || "\u2014",
          quantite: f.Quantite,
          prix: f.Prix,
          date: f.DateTransaction,
          compte: f.CompteUtilise,
          commentaire: f.Commentaire,
          isFound: f.Commentaire && f.Commentaire.startsWith("Restreint"),
          dateDeclaration: item.createdDateTime,
          id: item.id
        };
      }).sort((a, b) => new Date(b.dateDeclaration) - new Date(a.dateDeclaration));

      if (query) items = items.filter(i => i.titre.toLowerCase().includes(query.toLowerCase()));
      historyCache = items; // Mise en cache
      hideElement("history-loading");
      if (items.length === 0) { showElement("history-empty"); return; }
      showElement("history-table-container");
      _renderHistoryRows(items);

    } catch (err) {
      console.error("Erreur chargement historique:", err);
      hideElement("history-loading");
      showElement("history-empty");
    }
  } else {
    const key = getHistoryKey();
    let items = JSON.parse(localStorage.getItem(key) || "[]");
    if (query) items = items.filter(i => i.titre.toLowerCase().includes(query.toLowerCase()));
    hideElement("history-loading");
    if (items.length === 0) { showElement("history-empty"); return; }
    showElement("history-table-container");
    _renderHistoryRows(items);
  }
}

function _renderHistoryRows(items) {
  const tbody = document.getElementById("history-tbody");
  tbody.innerHTML = "";
  items.forEach(item => {
    const dateDecl = new Date(item.dateDeclaration).toLocaleDateString("fr-FR");
    const heureDecl = new Date(item.dateDeclaration).toLocaleTimeString("fr-FR");

    if (item.type_entree === "consultation") {
      const statusBadge = item.isFound
        ? `<span class="badge badge-restricted">Restreint</span>`
        : `<span class="badge badge-active">Non restreint</span>`;
      tbody.innerHTML += `
        <tr>
          <td><strong>${item.titre}</strong></td>
          <td><span class="badge badge-nda">Consultation</span></td>
          <td colspan="3">${statusBadge}</td>
          <td>${dateDecl} ${heureDecl}</td>
          <td><button class="btn-pdf" onclick='generateRestrictedPDFFromHistory(${JSON.stringify(item).replace(/'/g,"&#39;")})'>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            PDF
          </button></td>
        </tr>`;
    } else {
      const dateTx = item.date ? new Date(item.date).toLocaleDateString("fr-FR") : "\u2014";
      tbody.innerHTML += `
        <tr>
          <td><strong>${item.titre}</strong></td>
          <td><span class="badge badge-active">D\u00e9claration</span></td>
          <td>${item.quantite || "\u2014"}</td>
          <td>${item.prix ? item.prix + " \u20ac" : "\u2014"}</td>
          <td>${dateTx}</td>
          <td>${dateDecl}</td>
          <td><button class="btn-pdf" onclick='regenPDF(${JSON.stringify(item).replace(/'/g,"&#39;")})'>
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            PDF
          </button></td>
        </tr>`;
    }
  });
}

function toggleTooltip() {
  const tooltip = document.getElementById("rules-tooltip");
  tooltip.classList.toggle("hidden");
  // Ferme si on clique ailleurs
  if (!tooltip.classList.contains("hidden")) {
    setTimeout(() => {
      document.addEventListener("click", function closeTooltip(e) {
        if (!document.querySelector(".tooltip-container").contains(e.target)) {
          tooltip.classList.add("hidden");
          document.removeEventListener("click", closeTooltip);
        }
      });
    }, 10);
  }
}

// -----------------------------------------------
// FORMULAIRE CADEAU
// -----------------------------------------------
document.getElementById("btn-submit-gift").addEventListener("click", async () => {
  const description = document.getElementById("g-description").value.trim();
  const emetteur = document.getElementById("g-emetteur").value.trim();
  const destinataire = document.getElementById("g-destinataire").value.trim();
  const date = document.getElementById("g-date").value;
  const valeur = document.getElementById("g-valeur").value;

  if (!description || !emetteur || !destinataire || !date) {
    showGiftMessage("Veuillez remplir tous les champs obligatoires (*)", "error");
    return;
  }

  const btn = document.getElementById("btn-submit-gift");
  btn.disabled = true;

  try {
    await callGraphAPI(
      `/sites/${siteIdHistory}/lists/${encodeURIComponent(CONFIG.sharepointHistory.listGifts)}/items`,
      "POST",
      { fields: {
        Title: description,
        Emetteur: emetteur,
        Destinataire: destinataire,
        Date: date ? new Date(date + "T00:00:00").toISOString() : null,
        Valeur: valeur ? parseFloat(valeur) : null
      }}
    );

    showGiftMessage("✓ Déclaration enregistrée ! Génération de l'attestation et ouverture du mail...", "success");

    // Génère le PDF
    generateGiftPDF({ description, emetteur, destinataire, date, valeur });

    // Ouvre Outlook vers CPL
    const userName = currentUser.displayName || currentUser.mail;
    const dateFr = date ? new Date(date).toLocaleDateString("fr-FR") : "—";
    const subject = encodeURIComponent(`Déclaration cadeau — ${description} — ${userName}`);
    const body = encodeURIComponent(
      `Bonjour,\n\nVeuillez trouver ci-joint mon attestation de déclaration de cadeau.\n\n` +
      `Déclarant : ${userName}\n` +
      `Description : ${description}\n` +
      `Émetteur : ${emetteur}\n` +
      `Destinataire : ${destinataire}\n` +
      `Date : ${dateFr}\n` +
      `Valeur estimée : ${valeur ? valeur + " €" : "Non renseignée"}\n\n` +
      `Cordialement,\n${userName}`
    );
    setTimeout(() => {
      window.location.href = `mailto:cpl@eiffel-ig.com?subject=${subject}&body=${body}`;
    }, 1000);

    setTimeout(() => {
      ["g-description","g-emetteur","g-destinataire","g-date","g-valeur"].forEach(id => {
        document.getElementById(id).value = "";
      });
      hideElement("gift-msg");
      document.getElementById("gift-form-container").classList.add("hidden");
    }, 5000);

  } catch (err) {
    console.error("Erreur enregistrement cadeau:", err);
    showGiftMessage("Erreur lors de l'enregistrement. Vérifiez vos permissions.", "error");
  } finally {
    btn.disabled = false;
  }
});

function generateGiftPDF(data) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const now = new Date();
  const dateStr = now.toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;
  const dateFr = data.date ? new Date(data.date).toLocaleDateString("fr-FR") : "—";

  // En-tête
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

  // Titre
  doc.setTextColor(0, 20, 59);
  doc.setFontSize(15);
  doc.setFont("helvetica", "bold");
  doc.text("ATTESTATION DE DÉCLARATION DE CADEAU", 105, 55, { align: "center" });
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(80, 80, 80);
  doc.text("Registre cadeaux — Eiffel Investment Group", 105, 63, { align: "center" });

  doc.setDrawColor(0, 102, 96);
  doc.setLineWidth(0.8);
  doc.line(20, 68, 190, 68);

  let y = 82;
  const addRow = (label, value) => {
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10.5);
    doc.setTextColor(0, 20, 59);
    doc.text(label, 20, y);
    doc.setFont("helvetica", "normal");
    doc.setTextColor(50, 50, 50);
    doc.text(String(value || "—"), 80, y);
    y += 9;
  };

  addRow("Déclarant :", userName);
  addRow("Email :", currentUser.mail || "—");
  addRow("Date de déclaration :", dateStr);
  addRow("Heure :", timeStr);

  y += 4;
  doc.setDrawColor(200, 200, 200);
  doc.line(20, y, 190, y);
  y += 10;

  doc.setFontSize(11);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(0, 20, 59);
  doc.text("DÉTAILS DU CADEAU", 20, y);
  y += 10;

  addRow("Description :", data.description);
  addRow("Émetteur :", data.emetteur);
  addRow("Destinataire :", data.destinataire);
  addRow("Date :", dateFr);
  addRow("Valeur estimée :", data.valeur ? `${data.valeur} €` : "Non renseignée");

  y += 6;
  doc.setFillColor(232, 244, 244);
  doc.rect(20, y - 4, 170, 32, "F");
  doc.setDrawColor(0, 102, 96);
  doc.setLineWidth(0.5);
  doc.rect(20, y - 4, 170, 32);
  doc.setFontSize(10);
  doc.setFont("helvetica", "normal");
  doc.setTextColor(50, 50, 50);
  [
    `Je soussigné(e), ${userName}, atteste avoir déclaré le cadeau`,
    `ci-dessus dans le cadre de mes obligations de conformité.`,
    `Cette déclaration a été enregistrée le ${dateStr} à ${timeStr}`,
    `dans le Registre cadeaux d'Eiffel Investment Group.`
  ].forEach(line => { doc.text(line, 28, y + 5); y += 8; });

  y += 14;
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
  doc.text("Document généré automatiquement — Eiffel Investment Group Compliance Portal", 105, 291, { align: "center" });

  doc.save(`Attestation_Cadeau_${now.toISOString().slice(0,10)}.pdf`);
}

function showGiftMessage(text, type) {
  const el = document.getElementById("gift-msg");
  el.textContent = text;
  el.className = `form-message ${type}`;
  showElement("gift-msg");
}

function toggleGiftForm() {
  const container = document.getElementById("gift-form-container");
  const isHidden = container.classList.contains("hidden");
  if (isHidden) {
    container.classList.remove("hidden");
    container.scrollIntoView({ behavior: "smooth", block: "start" });
  } else {
    container.classList.add("hidden");
  }
}

function toggleTransactionForm() {
  const container = document.getElementById("transaction-form-container");
  const isHidden = container.classList.contains("hidden");
  if (isHidden) {
    container.classList.remove("hidden");
    container.scrollIntoView({ behavior: "smooth", block: "start" });
  } else {
    container.classList.add("hidden");
  }
}

function regenPDF(item) { generateDeclarationPDF(item); }

function generateRestrictedPDFFromHistory(item) {
  // Restaure temporairement les données pour regénérer le PDF
  const savedItems = allRestrictedItems;
  if (item.foundItems) allRestrictedItems = item.foundItems;
  generateRestrictedPDF(item.titre);
  allRestrictedItems = savedItems;
}

async function callGraphAPI(endpoint, method = "GET", body = null, extraHeaders = {}) {
  let tokenResponse;
  try {
    tokenResponse = await msalInstance.acquireTokenSilent({ ...graphScopes, account: msalInstance.getActiveAccount() });
  } catch {
    await msalInstance.acquireTokenRedirect(graphScopes);
    return;
  }
  const options = {
    method,
    headers: {
      "Authorization": `Bearer ${tokenResponse.accessToken}`,
      "Content-Type": "application/json",
      ...extraHeaders
    }
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
