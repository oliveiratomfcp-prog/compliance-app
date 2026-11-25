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
    listGifts: "Registre cadeaux",
    listPositions: "positions portefeuilles",
    listPositionsHistory: "Historique positions",
    libPositionsBackup: "Sauvegardes positions"
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const graphScopes = { scopes: ["User.Read", "Sites.Read.All", "Sites.ReadWrite.All"] };

let currentUser = null;
let siteId = null;
let siteIdHistory = null;
let allRestrictedItems = [];
let isCPL = false;

// Utilisé par translations.js pour conserver le libellé "Compliance" après changement de langue
function isCurrentUserCPL() { return isCPL; }

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
  isCPL = isCPLUser(currentUser.mail) || isCPLUser(currentUser.userPrincipalName);
  updateRoleLabel();
  if (isCPL) insertPositionsDepositUI();

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

// silent : rafraîchissement automatique. Pas d'écran de chargement, la recherche et le tri en
// cours sont conservés, et un échec ne remplace jamais les données déjà affichées.
let restrictedLoading = false;
let restrictedLoadedAt = 0;

async function loadRestrictedList(silent) {
  if (restrictedLoading) return;
  restrictedLoading = true;
  try {
    await loadRestrictedListInner(!!silent);
  } finally {
    restrictedLoading = false;
  }
}

async function loadRestrictedListInner(silent) {
  if (!silent) {
    showElement("restricted-loading");
    hideElement("restricted-table-container");
    hideElement("restricted-error");
  }

  // Positions en portefeuille : chargement séparé qui ne rejette jamais, pour ne
  // jamais empêcher l'affichage des NDA et informations privilégiées.
  const positionsPromise = loadPortfolioPositionsSafe();

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

    let positions = await positionsPromise;
    // Rafraîchissement automatique : une lecture ratée des positions ne doit pas faire
    // disparaître celles déjà affichées (on garde la dernière lecture réussie).
    if (silent && positions.status !== "ok" && positionsState.status === "ok") {
      console.warn("Rafraîchissement des positions échoué, dernière lecture conservée :", positions.reason);
      positions = positionsState;
    } else {
      applyPositionsState(positions);
    }

    allRestrictedItems = [...ndaItems, ...infoPrivItems, ...positions.items];
    restrictedLoadedAt = Date.now();
    hideElement("restricted-loading");
    if (silent) refreshRestrictedView();
    else renderRestrictedTable(allRestrictedItems);
    showElement("restricted-table-container");

  } catch (err) {
    console.error("Erreur chargement:", err);
    if (silent) return;
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

  const esc = PositionsCore.escapeHtml;
  const fragment = document.createDocumentFragment();
  items.forEach(item => {
    const badgeClass = item.source === "NDA List" ? "badge-nda"
      : item.source === PositionsCore.POSITION_SOURCE ? "badge-portefeuille"
      : "badge-restricted";
    const dateFin = item.dateFin;
    // "En cours" : uniquement pour les informations privil\u00e9gi\u00e9es dont la date de fin est future
    const isFutureEnd = item.source === "Info Priv" && dateFin && dateFin !== "\u2014" && new Date(dateFin) > new Date();
    const dateFinCell = isFutureEnd
      ? `<span class="pill-live"><span class="live-dot"></span>En cours</span>`
      : esc(formatDate(dateFin));
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><strong>${esc(item.nom)}</strong></td>
      <td style="font-family:monospace;font-size:13px">${esc(item.isin)}</td>
      <td><span class="badge ${badgeClass}">${esc(item.type)}</span></td>
      <td>${item.equipe && item.equipe !== "\u2014" ? `<span class="badge badge-active">${esc(item.equipe)}</span>` : "\u2014"}</td>
      <td>${item.signataire !== "\u2014" ? esc(item.signataire) : "\u2014"}</td>
      <td>${esc(formatDate(item.dateDebut))}</td>
      <td>${dateFinCell}</td>`;
    fragment.appendChild(tr);
  });
  tbody.appendChild(fragment);
}

let currentSort = { col: null, dir: 1 };

// Réaffiche le tableau avec la recherche et le tri en cours (après un rafraîchissement)
function refreshRestrictedView() {
  const query = document.getElementById("search-restricted").value.toLowerCase().trim();
  const items = query
    ? allRestrictedItems.filter(item => PositionsCore.matchesSearch(item, query))
    : [...allRestrictedItems];
  const col = currentSort.col;
  if (col) {
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
  }
  renderRestrictedTable(items);
}

// Rafraîchissement automatique toutes les 2 minutes : reporté si l'onglet est masqué ou si un
// dépôt de positions est en cours. Vérification toutes les 15 s, chargement au plus toutes les 2 min.
const RESTRICTED_REFRESH_MS = 2 * 60 * 1000;

function autoRefreshRestricted() {
  if (!siteId || restrictedLoading || positionsDepositBusy || document.visibilityState !== "visible") return;
  if (Date.now() - restrictedLoadedAt < RESTRICTED_REFRESH_MS - 5000) return;
  loadRestrictedList(true);
}

setInterval(autoRefreshRestricted, 15000);
document.addEventListener("visibilitychange", autoRefreshRestricted);

const RESTRICTED_SORT_COLUMNS = ["nom", "isin", "type", "equipe", "signataire", "dateDebut", "dateFin"];

function sortTable(col) {
  if (currentSort.col === col) {
    currentSort.dir *= -1;
  } else {
    currentSort.col = col;
    currentSort.dir = 1;
  }

  RESTRICTED_SORT_COLUMNS.forEach(c => {
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
    ? allRestrictedItems.filter(item => PositionsCore.matchesSearch(item, query))
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
    RESTRICTED_SORT_COLUMNS.forEach(c => {
      const el = document.getElementById(`sort-${c}`);
      if (el) { el.textContent = "\u2195"; el.parentElement.classList.remove("sort-asc", "sort-desc"); }
    });
    const filtered = query
      ? allRestrictedItems.filter(item => PositionsCore.matchesSearch(item, query))
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
    ? allRestrictedItems.filter(i => PositionsCore.matchesSearch(i, query))
    : allRestrictedItems;

  const bom = "\uFEFF";
  const csvCell = v => `"${String(v === undefined || v === null ? "" : v).replace(/"/g, '""')}"`;
  const headers = ["Soci\u00e9t\u00e9 / Titre", "Code ISIN", "Type de restriction", "\u00c9quipe", "Date de d\u00e9but", "Date de fin"];
  const rows = items.map(item => [
    csvCell(item.nom), csvCell(item.isin), csvCell(item.type),
    csvCell(item.equipe || ""), csvCell(formatDate(item.dateDebut)), csvCell(formatDate(item.dateFin))
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
// Statut d'un résultat de recherche : "restricted" si au moins une vraie restriction,
// "portfolio" si uniquement des positions en portefeuille (non bloquant), "clean" sinon.
function matchStatus(matches) {
  if (!matches || !matches.length) return "clean";
  return matches.every(m => PositionsCore.isPositionItem(m)) ? "portfolio" : "restricted";
}

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
  let portfolio = 0;
  let clean = 0;
  const resultData = [];

  lastBulkPositionsUnavailable = positionsUnavailable();
  if (lastBulkPositionsUnavailable) {
    const warn = document.createElement("div");
    warn.className = "positions-banner";
    warn.textContent = t("pos_bulk_unavailable");
    resultsEl.appendChild(warn);
  }

  entries.forEach(entry => {
    const matches = allRestrictedItems.filter(item => PositionsCore.matchesBulkEntry(item, entry));
    // Rouge : au moins une vraie restriction (NDA, information privil\u00E9gi\u00E9e) ;
    // jaune : uniquement des positions en portefeuille (information, non bloquant).
    const status = matchStatus(matches);
    const isRestricted = status === "restricted";
    const isPortfolio = status === "portfolio";
    if (isRestricted) restricted++; else if (isPortfolio) portfolio++; else clean++;
    resultData.push({ entry, isRestricted, isPortfolio, matches });

    const div = document.createElement("div");
    div.className = `bulk-result-item ${isRestricted ? "restricted" : isPortfolio ? "portfolio" : "clean"}`;
    const detail = PositionsCore.escapeHtml(PositionsCore.describeMatchTypes(matches));
    div.innerHTML = `
      <span class="bulk-result-icon">${isRestricted ? "\uD83D\uDEAB" : isPortfolio ? "\uD83D\uDCBC" : "\u2705"}</span>
      <div style="flex:1">
        <strong>${PositionsCore.escapeHtml(entry)}</strong>
        ${isRestricted
          ? `<span style="margin-left:8px;font-size:12px;color:#c53030">Restreint \u2014 ${detail}</span>`
          : isPortfolio
            ? `<span style="margin-left:8px;font-size:12px;color:#92400e">${detail} \u2014 non bloquant</span>`
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
      ${portfolio ? `<span style="color:#92400e"><strong>${portfolio}</strong> en portefeuille</span>,` : ""}
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
  const portf = resultData.filter(r => r.isPortfolio).length;
  doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59);
  doc.text("R\u00e9sum\u00e9 :", 20, y); doc.setFont("helvetica", "normal"); doc.setTextColor(50,50,50);
  doc.text(`${resultData.length} titre(s) \u2014 ${restr} restreint(s), ${portf ? `${portf} en portefeuille, ` : ""}${resultData.length - restr - portf} autoris\u00e9(s)`, 80, y); y += 14;

  if (lastBulkPositionsUnavailable) {
    y = addPositionsUnavailablePdfNotice(doc, y);
  } else {
    y = addPositionsFreshnessPdfLine(doc, y - 6) + 6;
  }

  doc.setDrawColor(200,200,200); doc.line(20, y, 190, y); y += 10;
  doc.setFontSize(11); doc.setFont("helvetica", "bold"); doc.setTextColor(0,20,59);
  doc.text("D\u00c9TAIL", 20, y); y += 10;

  resultData.forEach(r => {
    if (y > 265) { doc.addPage(); y = 20; }
    if (r.isRestricted) {
      doc.setFont("helvetica", "normal"); doc.setFontSize(9);
      const detailLines = doc.splitTextToSize(`Restreint \u2014 ${PositionsCore.describeMatchTypes(r.matches)}`, 160);
      const boxH = 12 + detailLines.length * 4;
      doc.setFillColor(253, 232, 232); doc.rect(20, y-4, 170, boxH, "F");
      doc.setDrawColor(245,198,198); doc.rect(20, y-4, 170, boxH);
      doc.setFont("helvetica", "bold"); doc.setTextColor(197,48,48); doc.setFontSize(10);
      doc.text(`\u26a0 ${r.entry}`, 25, y+4);
      doc.setFont("helvetica", "normal"); doc.setFontSize(9);
      doc.text(detailLines, 25, y+11);
      y += boxH + 4;
    } else if (r.isPortfolio) {
      doc.setFont("helvetica", "normal"); doc.setFontSize(9);
      const detailLines = doc.splitTextToSize(`${PositionsCore.describeMatchTypes(r.matches)} — non bloquant`, 160);
      const boxH = 12 + detailLines.length * 4;
      doc.setFillColor(254, 243, 199); doc.rect(20, y-4, 170, boxH, "F");
      doc.setDrawColor(253, 230, 138); doc.rect(20, y-4, 170, boxH);
      doc.setFont("helvetica", "bold"); doc.setTextColor(146, 64, 14); doc.setFontSize(10);
      doc.text(`• ${r.entry}`, 25, y+4);
      doc.setFont("helvetica", "normal"); doc.setFontSize(9);
      doc.text(detailLines, 25, y+11);
      y += boxH + 4;
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
    restr + portf > 0,
    resultData.filter(r => r.isRestricted || r.isPortfolio).flatMap(r => r.matches)
  );

  doc.save(`Rapport_Verification_${now.toISOString().slice(0,10)}.pdf`);
}

document.getElementById("btn-generate-restricted-pdf").addEventListener("click", () => {
  if (!currentUser) return;
  const query = document.getElementById("search-restricted").value.trim();
  generateRestrictedPDF(query);
});

function generateRestrictedPDF(query, opts = {}) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const now = new Date();
  const dateStr = now.toLocaleDateString("fr-FR", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
  const timeStr = now.toLocaleTimeString("fr-FR");
  const userName = currentUser.displayName || currentUser.mail;

  // Cherche si le titre est dans la restricted list
  const found = query
    ? allRestrictedItems.filter(item => PositionsCore.matchesSecurity(item, query))
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
  if (isFound && matchStatus(found) === "portfolio") {
    // Uniquement des positions en portefeuille : information, non bloquant (jaune)
    doc.setFontSize(10);
    const lines = doc.splitTextToSize(`Le titre "${searchLabel}" est d\u00e9tenu en portefeuille par Eiffel Investment Group. Il n'est pas sous restriction.`, 155);
    const boxH = 17 + lines.length * 5;
    doc.setFillColor(254, 243, 199);
    doc.rect(20, y - 4, 170, boxH, "F");
    doc.setDrawColor(217, 119, 6);
    doc.setLineWidth(0.8);
    doc.rect(20, y - 4, 170, boxH);
    doc.setFontSize(11);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(146, 64, 14);
    doc.text("TITRE EN PORTEFEUILLE (NON BLOQUANT)", 28, y + 5);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(10);
    doc.text(lines, 28, y + 14);
    y += boxH + 8;
  } else if (isFound) {
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
  }
  if (isFound) {

    // Détail des titres trouvés
    doc.setFontSize(10.5);
    doc.setFont("helvetica", "bold");
    doc.setTextColor(0, 20, 59);
    doc.text("D\u00e9tail des titres trouv\u00e9s :", 20, y);
    y += 8;

    found.forEach(item => {
      if (y > 262) { doc.addPage(); y = 20; }
      doc.setFont("helvetica", "normal");
      doc.setTextColor(50, 50, 50);
      doc.setFontSize(10);
      doc.text(`\u2022 ${item.nom}${item.isin !== "\u2014" ? ` (ISIN: ${item.isin})` : ""} \u2014 ${item.type}`, 24, y);
      if (PositionsCore.isPositionItem(item)) {
        doc.text(`  Strat\u00e9gie : ${item.equipe || "\u2014"}`, 24, y + 6);
      } else {
        doc.text(`  Date de fin : ${formatDate(item.dateFin)}`, 24, y + 6);
      }
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

  if (!opts.fromHistory && positionsUnavailable()) {
    if (y > 230) { doc.addPage(); y = 20; }
    y = addPositionsUnavailablePdfNotice(doc, y);
  } else if (!opts.fromHistory) {
    y = addPositionsFreshnessPdfLine(doc, y - 2) + 2;
  }

  if (y > 220) { doc.addPage(); y = 20; }
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
  const restricted = allRestrictedItems.filter(item => PositionsCore.matchesSecurity(item, titre));

  if (restricted.length > 0) {
    const onlyPositions = restricted.every(item => PositionsCore.isPositionItem(item));
    showFormMessage(
      onlyPositions
        ? `\u26a0 ${t("tx_portfolio_restricted")}`
        : `\u26a0 "${titre}" figure sur la Restricted List \u2014 cette transaction ne peut pas \u00eatre demand\u00e9e. Contactez le service Compliance.`,
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
          Commentaire: !isFound ? "Non restreint"
            : (matchStatus(foundItems) === "portfolio" ? "En portefeuille (non bloquant) - " : "Restreint - ") + foundItems.map(i => i.nom).join(", ")
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
          isFound: !!f.Commentaire && (f.Commentaire.startsWith("Restreint") || f.Commentaire.startsWith("En portefeuille")),
          isPortfolio: !!f.Commentaire && f.Commentaire.startsWith("En portefeuille"),
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
      const portfolioOnly = item.isPortfolio || (item.isFound && item.foundItems && matchStatus(item.foundItems) === "portfolio");
      const statusBadge = portfolioOnly
        ? `<span class="badge badge-portefeuille-info">En portefeuille (non bloquant)</span>`
        : item.isFound
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

// Résout dynamiquement les noms internes des colonnes SharePoint à partir de leurs libellés
// affichés (les noms internes générés par SharePoint diffèrent des libellés dès qu'ils
// contiennent espaces, accents ou tirets). Le résultat est mis en cache pour la session.
let giftColumnMap = null;
async function getGiftColumnMap() {
  if (giftColumnMap) return giftColumnMap;

  const REQUIRED_DISPLAY_NAMES = [
    "Type de cadeau",
    "Nature du tiers",
    "Nature du tiers - Précision",
    "Opération en cours",
    "Sort du cadeau",
    "Sort du cadeau - Précision"
  ];

  const res = await callGraphAPI(
    `/sites/${siteIdHistory}/lists/${encodeURIComponent(CONFIG.sharepointHistory.listGifts)}/columns?$select=name,displayName`
  );
  const map = {};
  (res?.value || []).forEach(c => { map[c.displayName] = c.name; });

  const missing = REQUIRED_DISPLAY_NAMES.filter(n => !map[n]);
  if (missing.length) {
    throw new Error(`Colonnes introuvables dans la liste "Registre cadeaux" : ${missing.join(", ")}. Vérifiez les libellés des colonnes SharePoint.`);
  }

  giftColumnMap = map;
  return giftColumnMap;
}

// Affiche/masque le champ de précision "Autre" pour la nature du tiers
document.getElementById("g-nature-tiers").addEventListener("change", (e) => {
  const group = document.getElementById("g-nature-tiers-precision-group");
  if (e.target.value === "Autre") {
    group.classList.remove("hidden");
  } else {
    group.classList.add("hidden");
    document.getElementById("g-nature-tiers-precision").value = "";
  }
});

// Affiche/masque le champ de précision "Autre" pour le sort du cadeau
document.getElementById("g-sort-cadeau").addEventListener("change", (e) => {
  const group = document.getElementById("g-sort-cadeau-precision-group");
  if (e.target.value === "Autre") {
    group.classList.remove("hidden");
  } else {
    group.classList.add("hidden");
    document.getElementById("g-sort-cadeau-precision").value = "";
  }
});

// Met en surbrillance l'option sélectionnée du toggle Oui/Non
document.querySelectorAll('input[name="g-operation-en-cours"]').forEach((radio) => {
  radio.addEventListener("change", () => {
    document.querySelectorAll('input[name="g-operation-en-cours"]').forEach((r) => {
      r.closest(".radio-toggle-option").classList.toggle("active", r.checked);
    });
  });
});

document.getElementById("btn-submit-gift").addEventListener("click", async () => {
  const typeCadeau = document.getElementById("g-type-cadeau").value;
  const description = document.getElementById("g-description").value.trim();
  const natureTiers = document.getElementById("g-nature-tiers").value;
  const natureTiersPrecision = document.getElementById("g-nature-tiers-precision").value.trim();
  const emetteur = document.getElementById("g-emetteur").value.trim();
  const destinataire = document.getElementById("g-destinataire").value.trim();
  const operationEnCoursEl = document.querySelector('input[name="g-operation-en-cours"]:checked');
  const operationEnCours = operationEnCoursEl ? operationEnCoursEl.value : "";
  const date = document.getElementById("g-date").value;
  const valeur = document.getElementById("g-valeur").value;
  const sortCadeau = document.getElementById("g-sort-cadeau").value;
  const sortCadeauPrecision = document.getElementById("g-sort-cadeau-precision").value.trim();

  const missingRequired = !typeCadeau || !description || !natureTiers || !emetteur ||
    !destinataire || !operationEnCours || !date || !sortCadeau;
  const missingNaturePrecision = natureTiers === "Autre" && !natureTiersPrecision;
  const missingSortPrecision = sortCadeau === "Autre" && !sortCadeauPrecision;

  if (missingRequired || missingNaturePrecision || missingSortPrecision) {
    showGiftMessage("Veuillez remplir tous les champs obligatoires (*)", "error");
    return;
  }

  const btn = document.getElementById("btn-submit-gift");
  btn.disabled = true;

  try {
    const col = await getGiftColumnMap();
    await callGraphAPI(
      `/sites/${siteIdHistory}/lists/${encodeURIComponent(CONFIG.sharepointHistory.listGifts)}/items`,
      "POST",
      { fields: {
        Title: description,
        [col["Type de cadeau"]]: typeCadeau,
        [col["Nature du tiers"]]: natureTiers,
        [col["Nature du tiers - Précision"]]: natureTiers === "Autre" ? natureTiersPrecision : null,
        Emetteur: emetteur,
        Destinataire: destinataire,
        [col["Opération en cours"]]: operationEnCours,
        Date: date ? new Date(date + "T00:00:00").toISOString() : null,
        Valeur: valeur ? parseFloat(valeur) : null,
        [col["Sort du cadeau"]]: sortCadeau,
        [col["Sort du cadeau - Précision"]]: sortCadeau === "Autre" ? sortCadeauPrecision : null
      }}
    );

    showGiftMessage("✓ Déclaration enregistrée ! Génération de l'attestation et ouverture du mail...", "success");

    // Génère le PDF
    generateGiftPDF({
      typeCadeau, description, natureTiers, natureTiersPrecision, emetteur, destinataire,
      operationEnCours, date, valeur, sortCadeau, sortCadeauPrecision
    });

    // Ouvre Outlook vers CPL
    const userName = currentUser.displayName || currentUser.mail;
    const dateFr = date ? new Date(date).toLocaleDateString("fr-FR") : "—";
    const natureTiersLabel = natureTiers === "Autre" ? `Autre (${natureTiersPrecision})` : natureTiers;
    const sortCadeauLabel = sortCadeau === "Autre" ? `Autre (${sortCadeauPrecision})` : sortCadeau;
    const subject = encodeURIComponent(`Déclaration cadeau — ${description} — ${userName}`);
    const body = encodeURIComponent(
      `Bonjour,\n\nVeuillez trouver ci-joint mon attestation de déclaration de cadeau.\n\n` +
      `Déclarant : ${userName}\n` +
      `Type de cadeau : ${typeCadeau}\n` +
      `Description : ${description}\n` +
      `Nature du tiers : ${natureTiersLabel}\n` +
      `Émetteur : ${emetteur}\n` +
      `Destinataire : ${destinataire}\n` +
      `Opération en cours avec ce tiers : ${operationEnCours}\n` +
      `Date : ${dateFr}\n` +
      `Valeur estimée : ${valeur ? valeur + " €" : "Non renseignée"}\n` +
      `Sort du cadeau : ${sortCadeauLabel}\n\n` +
      `Cordialement,\n${userName}`
    );
    setTimeout(() => {
      window.location.href = `mailto:cpl@eiffel-ig.com?subject=${subject}&body=${body}`;
    }, 1000);

    setTimeout(() => {
      ["g-type-cadeau","g-description","g-nature-tiers","g-nature-tiers-precision","g-emetteur",
       "g-destinataire","g-date","g-valeur","g-sort-cadeau","g-sort-cadeau-precision"].forEach(id => {
        document.getElementById(id).value = "";
      });
      document.querySelectorAll('input[name="g-operation-en-cours"]').forEach((r) => {
        r.checked = false;
        r.closest(".radio-toggle-option").classList.remove("active");
      });
      document.getElementById("g-nature-tiers-precision-group").classList.add("hidden");
      document.getElementById("g-sort-cadeau-precision-group").classList.add("hidden");
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
    y += 7;
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

  const natureTiersLabel = data.natureTiers === "Autre" ? `Autre (${data.natureTiersPrecision})` : data.natureTiers;
  const sortCadeauLabel = data.sortCadeau === "Autre" ? `Autre (${data.sortCadeauPrecision})` : data.sortCadeau;

  addRow("Type de cadeau :", data.typeCadeau);
  addRow("Description :", data.description);
  addRow("Nature du tiers :", natureTiersLabel);
  addRow("Émetteur :", data.emetteur);
  addRow("Destinataire :", data.destinataire);
  addRow("Opération en cours :", data.operationEnCours);
  addRow("Date :", dateFr);
  addRow("Valeur estimée :", data.valeur ? `${data.valeur} €` : "Non renseignée");
  addRow("Sort du cadeau :", sortCadeauLabel);

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
  generateRestrictedPDF(item.titre, { fromHistory: true });
  allRestrictedItems = savedItems;
}

// -----------------------------------------------
// POSITIONS EN PORTEFEUILLE
// -----------------------------------------------
// Troisième source de la Restricted List (liste "positions portefeuilles" sur
// CPLDashboard). La logique testable (parsing, aperçu, garde de suppression,
// remplacement) est dans positions-core.js ; ce bloc ne gère que Graph et le DOM.

const POSITIONS_LOAD_TIMEOUT_MS = 30000;
// Au-delà de ce nombre de jours sans nouveau dépôt, les positions sont signalées comme anciennes
const POSITIONS_MAX_AGE_DAYS = 7;
const POSITIONS_REQUIRED_COLUMNS = ["Title", "ISIN", "Fonds", "IdDepot"];
const POSITIONS_TEXT_COLUMNS = ["ISIN", "Fonds", "IdDepot"];

let positionsTarget = null;        // cible figée, résolue une seule fois (PositionsCore.makePositionsTarget)
let positionsTargetError = null;   // message si la liste n'a pas pu être résolue
let positionsColumnsOk = false;
let positionsColumnsError = null;
let positionsState = { status: "loading", reason: "", items: [], rawItems: [], idDepots: [] };
let lastBulkPositionsUnavailable = false;
let positionsPreview = null;
let positionsDepositBusy = false;

function positionsUnavailable() { return positionsState.status === "unavailable"; }

// Historique des dépôts (liste) et sauvegardes de l'ancienne liste (bibliothèque), sur CPLDashboard.
// Facultatifs : s'ils manquent, le dépôt reste possible mais sans verrou ni sauvegarde, et la
// Compliance en est avertie. La liste doit contenir au moins la colonne Statut.
const positionsHistory = { resolved: false, listId: null, driveId: null, columns: new Set(), error: null, entries: [] };

class PositionsLockError extends Error {
  constructor(message) { super(message); this.name = "PositionsLockError"; }
}

// Remplace les {cle} d'un texte traduit
function tf(key, vars = {}) {
  return String(t(key)).replace(/\{(\w+)\}/g, (m, k) => (vars[k] !== undefined ? String(vars[k]) : m));
}

// Jeton sans redirection : une redirection en pleine opération de dépôt serait dangereuse
async function acquireGraphTokenNoRedirect() {
  const res = await msalInstance.acquireTokenSilent({ ...graphScopes, account: msalInstance.getActiveAccount() });
  return res.accessToken;
}

async function graphGetNoRedirect(endpoint) {
  const token = await acquireGraphTokenNoRedirect();
  const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    headers: { "Authorization": `Bearer ${token}` }
  });
  if (!response.ok) { const err = await response.json().catch(() => ({})); throw new Error(err.error?.message || `Erreur API: ${response.status}`); }
  return response.json();
}

async function graphGetAllPages(endpoint) {
  let items = [];
  let url = endpoint;
  while (url) {
    const data = await graphGetNoRedirect(url);
    items = items.concat(data.value || []);
    url = data["@odata.nextLink"]
      ? data["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
      : null;
  }
  return items;
}

// Résout une seule fois l'identifiant de la liste par son nom d'affichage exact
async function resolvePositionsTarget() {
  if (positionsTarget) return positionsTarget;
  if (!siteIdHistory) throw new Error("Site SharePoint CPLDashboard inaccessible.");
  const name = CONFIG.sharepointHistory.listPositions;
  const lists = await graphGetAllPages(`/sites/${siteIdHistory}/lists?$select=id,displayName`);
  const matches = lists.filter(l => l.displayName === name);
  if (matches.length === 0) throw new Error(`Liste "${name}" introuvable sur CPLDashboard.`);
  if (matches.length > 1) throw new Error(`Plusieurs listes nommées "${name}" sur CPLDashboard.`);
  const protectedNames = [
    CONFIG.sharepoint.listNDA, CONFIG.sharepoint.listInfoPriv,
    CONFIG.sharepointHistory.listHistory, CONFIG.sharepointHistory.listGifts
  ];
  positionsTarget = PositionsCore.makePositionsTarget({
    siteId: siteIdHistory,
    listId: matches[0].id,
    displayName: matches[0].displayName,
    forbiddenSiteIds: [siteId],
    forbiddenListIds: lists.filter(l => protectedNames.includes(l.displayName)).map(l => l.id)
  });
  return positionsTarget;
}

// Vérifie les noms internes des colonnes (et leur type texte) avant d'autoriser le dépôt
async function checkPositionsColumns(target) {
  try {
    const cols = await graphGetAllPages(`/sites/${target.siteId}/lists/${target.listId}/columns`);
    const byName = new Map(cols.map(c => [c.name, c]));
    const missing = POSITIONS_REQUIRED_COLUMNS.filter(n => !byName.has(n));
    const notText = POSITIONS_TEXT_COLUMNS.filter(n => byName.has(n) && !byName.get(n).text);
    const problems = [];
    if (missing.length) problems.push(`colonnes absentes : ${missing.join(", ")}`);
    if (notText.length) problems.push(`colonnes qui ne sont pas de type texte : ${notText.join(", ")}`);
    positionsColumnsOk = problems.length === 0;
    positionsColumnsError = problems.length ? problems.join(" ; ") : null;
  } catch (err) {
    positionsColumnsOk = false;
    positionsColumnsError = `lecture des colonnes impossible (${err.message})`;
  }
}

async function listPositionItems(target) {
  const expand = positionsColumnsOk ? "fields($select=Title,ISIN,Fonds,IdDepot)" : "fields";
  return graphGetAllPages(`/sites/${target.siteId}/lists/${target.listId}/items?$expand=${expand}&$top=500`);
}

// Ne rejette jamais : renseigne positionsHistory (liste, bibliothèque, colonnes, entrées récentes)
async function resolvePositionsHistory() {
  try {
    if (!positionsHistory.resolved) {
      const lists = await graphGetAllPages(`/sites/${siteIdHistory}/lists?$select=id,displayName`);
      const hist = lists.filter(l => l.displayName === CONFIG.sharepointHistory.listPositionsHistory);
      positionsHistory.listId = hist.length === 1 ? hist[0].id : null;
      const drives = await graphGetAllPages(`/sites/${siteIdHistory}/drives?$select=id,name`);
      const lib = drives.filter(d => d.name === CONFIG.sharepointHistory.libPositionsBackup);
      positionsHistory.driveId = lib.length === 1 ? lib[0].id : null;
      if (positionsHistory.listId) {
        const cols = await graphGetAllPages(`/sites/${siteIdHistory}/lists/${positionsHistory.listId}/columns`);
        positionsHistory.columns = new Set(cols.map(c => c.name));
        if (!positionsHistory.columns.has("Statut")) positionsHistory.listId = null;
      }
      positionsHistory.resolved = true;
    }
    positionsHistory.error = null;
    positionsHistory.entries = positionsHistory.listId ? await loadPositionsHistoryEntries() : [];
  } catch (err) {
    console.warn("Historique des dépôts de positions indisponible :", err);
    positionsHistory.error = err.message || String(err);
    positionsHistory.entries = [];
  }
}

async function loadPositionsHistoryEntries() {
  const items = await graphGetAllPages(`/sites/${siteIdHistory}/lists/${positionsHistory.listId}/items?$expand=fields&$top=500`);
  return items.map(it => {
    const f = it.fields || {};
    return {
      id: Number(it.id), title: f.Title || "", statut: f.Statut || "", auteur: f.Auteur || "",
      createdMs: it.createdDateTime ? Date.parse(it.createdDateTime) : NaN
    };
  });
}

function positionsHistoryHasColumn(name) { return positionsHistory.columns.has(name); }

// Écritures sans redirection (une redirection en plein dépôt serait dangereuse)
async function graphWriteNoRedirect(method, endpoint, body) {
  const token = await acquireGraphTokenNoRedirect();
  const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    method,
    headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body)
  });
  if (!response.ok) { const err = await response.json().catch(() => ({})); throw new Error(err.error?.message || `Erreur API: ${response.status}`); }
  const text = await response.text();
  return text ? JSON.parse(text) : null;
}

function positionsHistoryFields(fields) {
  const out = {};
  Object.keys(fields).forEach(k => { if (k === "Title" || positionsHistoryHasColumn(k)) out[k] = fields[k]; });
  return out;
}

// "Nom (email)" enregistré dans l'historique pour la traçabilité ; seul le nom est affiché
function displayAuthor(author) { return String(author || "").replace(/\s*\([^)]*\)\s*$/, "").trim(); }

function shortTime(ms) { return new Date(ms).toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" }); }

// Verrou optimiste : une entrée "En cours" est créée dans l'historique, puis relue. S'il existe
// une autre entrée "En cours" non expirée plus ancienne (id plus bas), ce dépôt s'efface devant
// elle. Renvoie l'entrée créée, ou null si l'historique n'est pas configuré (pas de verrou).
async function acquirePositionsLock(idDepot, preview, files) {
  if (!positionsHistory.listId) return null;
  const busy = lock => new PositionsLockError(tf("pos_lock_busy", { who: displayAuthor(lock.auteur) || "?", time: shortTime(lock.createdMs) }));

  const before = PositionsCore.activeLocks(await loadPositionsHistoryEntries(), Date.now());
  if (before.length) throw busy(before[0]);

  const author = `${currentUser.displayName || ""} (${currentUser.mail || currentUser.userPrincipalName || ""})`.trim();
  const created = await graphWriteNoRedirect("POST", `/sites/${siteIdHistory}/lists/${positionsHistory.listId}/items`, {
    fields: positionsHistoryFields({
      Title: idDepot, Statut: PositionsCore.LOCK_STATUS, Auteur: author,
      Fichiers: files.join(", ").slice(0, 250), NbLignes: preview.total, NbAnciennes: preview.currentTotal
    })
  });
  const ownId = Number(created.id);

  const after = PositionsCore.activeLocks(await loadPositionsHistoryEntries(), Date.now());
  if (after.length && after[0].id !== ownId) {
    try { await graphWriteNoRedirect("PATCH", `/sites/${siteIdHistory}/lists/${positionsHistory.listId}/items/${created.id}/fields`, positionsHistoryFields({ Statut: "Annulé", Detail: "Dépôt simultané" })); }
    catch (e) { console.warn("Entrée d'historique non annulée :", e); }
    throw busy(after[0]);
  }
  return { id: created.id };
}

// Sauvegarde de l'ancienne liste en CSV (avant sa suppression). Si la bibliothèque est configurée
// et que l'enregistrement échoue, l'erreur remonte : rien n'est supprimé.
let lastBackupFile = null;
async function backupPositionsBeforeDelete(idDepot, oldItems) {
  lastBackupFile = null;
  if (!positionsHistory.driveId) return;
  const name = PositionsCore.backupFileName(idDepot);
  const token = await acquireGraphTokenNoRedirect();
  const response = await fetch(`https://graph.microsoft.com/v1.0/drives/${positionsHistory.driveId}/root:/${encodeURIComponent(name)}:/content`, {
    method: "PUT",
    headers: { "Authorization": `Bearer ${token}`, "Content-Type": "text/csv" },
    body: new Blob([PositionsCore.buildBackupCsv(oldItems)], { type: "text/csv;charset=utf-8" })
  });
  if (!response.ok) { const err = await response.json().catch(() => ({})); throw new Error(err.error?.message || `HTTP ${response.status}`); }
  lastBackupFile = name;
}

// Clôture l'entrée d'historique (jamais bloquant). Renvoie un message si elle n'a pas pu être mise à jour.
async function finalizePositionsHistory(lock, result, fatal, preview) {
  if (!lock) return null;
  let statut = "Échec";
  let detail = "";
  if (fatal) detail = fatal.message;
  else if (result && result.ok) { statut = result.deleteFailed ? "Terminé avec erreurs" : "Terminé"; if (result.deleteFailed) detail = `${result.deleteFailed} suppression(s) échouée(s)`; }
  else if (result) detail = result.failure ? result.failure.message : "";
  const sum = list => list.reduce((n, x) => n + x.count, 0);
  try {
    await graphWriteNoRedirect("PATCH", `/sites/${siteIdHistory}/lists/${positionsHistory.listId}/items/${lock.id}/fields`, positionsHistoryFields({
      Statut: statut, NbAjoutees: sum(preview.added), NbRetirees: sum(preview.removed), NbManuelles: preview.manualRows.length,
      Sauvegarde: lastBackupFile || "", Detail: detail.slice(0, 250)
    }));
    return null;
  } catch (err) {
    console.warn("Historique du dépôt non mis à jour :", err);
    return tf("pos_history_not_updated", { n: Math.round(PositionsCore.LOCK_TTL_MS / 60000) });
  }
}

// Dernier dépôt présent dans la liste : date (portée par l'IdDepot) et auteur (historique, s'il existe)
function positionsDepotInfo() {
  if (positionsState.status !== "ok") return null;
  const latest = PositionsCore.latestDepot(positionsState.idDepots);
  if (!latest) return null;
  const entry = positionsHistory.entries.find(e => e.title === latest.idDepot && e.auteur);
  return { date: latest.date, author: entry ? displayAuthor(entry.auteur) : "", ageDays: (Date.now() - latest.date.getTime()) / 86400000 };
}

function formatDepotDate(d) {
  return `${d.toLocaleDateString("fr-FR")} ${d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" })}`;
}

function withTimeout(promise, ms, message) {
  let timer;
  const timeout = new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(message)), ms); });
  return Promise.race([promise, timeout]).finally(() => clearTimeout(timer));
}

// Ne rejette jamais : renvoie un état "ok" ou "unavailable"
async function loadPortfolioPositionsSafe() {
  try {
    const raw = await withTimeout((async () => {
      if (!siteIdHistory) throw new Error("Site SharePoint CPLDashboard inaccessible.");
      let target;
      try {
        target = await resolvePositionsTarget();
        positionsTargetError = null;
      } catch (err) {
        positionsTargetError = err.message;
        throw err;
      }
      if (!positionsColumnsOk) await checkPositionsColumns(target);
      await resolvePositionsHistory();
      return listPositionItems(target);
    })(), POSITIONS_LOAD_TIMEOUT_MS, "Délai de chargement dépassé (30 s).");
    return {
      status: "ok",
      reason: "",
      rawItems: raw,
      items: raw.map(PositionsCore.positionToRestrictedItem),
      idDepots: PositionsCore.distinctIdDepots(raw)
    };
  } catch (err) {
    console.error("Erreur chargement positions en portefeuille:", err);
    return { status: "unavailable", reason: err && err.message ? err.message : String(err), rawItems: [], items: [], idDepots: [] };
  }
}

function applyPositionsState(state) {
  positionsState = state;
  renderPositionsBanners();
  refreshPositionsPanelStatus();
}

// Bandeau neutre pour tous si les positions sont indisponibles ; détails techniques
// et avertissements (colonnes, dépôts multiples) réservés à la Compliance.
function renderPositionsBanners() {
  const banner = document.getElementById("positions-banner");
  const cplWarn = document.getElementById("positions-cpl-warning");
  const txBanner = document.getElementById("positions-tx-banner");

  if (banner) {
    banner.textContent = "";
    if (positionsUnavailable()) {
      const main = document.createElement("div");
      main.textContent = t("pos_unavailable_banner");
      banner.appendChild(main);
      if (isCPL && positionsState.reason) {
        const detail = document.createElement("div");
        detail.className = "positions-banner-detail";
        detail.textContent = tf("pos_unavailable_detail", { reason: positionsState.reason });
        banner.appendChild(detail);
      }
      banner.classList.remove("hidden");
    } else {
      banner.classList.add("hidden");
    }
  }

  if (txBanner) {
    txBanner.textContent = t("pos_unavailable_banner");
    txBanner.classList.toggle("hidden", !positionsUnavailable());
  }

  const info = positionsDepotInfo();
  const infoEl = document.getElementById("restricted-positions-info");
  if (infoEl) {
    infoEl.textContent = !info ? ""
      : info.author ? tf("pos_last_deposit_by", { date: formatDepotDate(info.date), author: info.author })
      : tf("pos_last_deposit", { date: formatDepotDate(info.date) });
  }

  // Positions anciennes ou liste vide : visible de tous (une vérification s'appuie sur ces données)
  const staleEl = document.getElementById("positions-stale-banner");
  if (staleEl) {
    let text = "";
    if (positionsState.status === "ok") {
      if (positionsState.items.length === 0) text = t("pos_empty_banner");
      else if (info && info.ageDays > POSITIONS_MAX_AGE_DAYS) {
        text = tf("pos_stale_banner", { date: formatDepotDate(info.date), days: Math.floor(info.ageDays) });
      }
    }
    staleEl.textContent = text;
    staleEl.classList.toggle("hidden", !text);
  }

  if (cplWarn) {
    cplWarn.textContent = "";
    const lines = [];
    if (isCPL && positionsState.status === "ok") {
      if (positionsColumnsError) lines.push(tf("pos_columns_warning", { detail: positionsColumnsError }));
      if (positionsState.idDepots.length > 1) lines.push(tf("pos_multi_depot", { n: positionsState.idDepots.length }));
      if (positionsHistory.error) lines.push(tf("pos_history_error", { detail: positionsHistory.error }));
      else if (positionsHistory.resolved && !positionsHistory.listId) lines.push(tf("pos_history_missing_list", { name: CONFIG.sharepointHistory.listPositionsHistory }));
      if (positionsHistory.resolved && !positionsHistory.driveId) lines.push(tf("pos_history_missing_lib", { name: CONFIG.sharepointHistory.libPositionsBackup }));
    }
    lines.forEach(text => {
      const div = document.createElement("div");
      div.textContent = text;
      cplWarn.appendChild(div);
    });
    cplWarn.classList.toggle("hidden", lines.length === 0);
  }
}

// Encadré ajouté aux PDF (attestation, vérification en masse) si les positions n'ont pas pu être vérifiées
function addPositionsUnavailablePdfNotice(doc, y) {
  doc.setFillColor(255, 243, 220);
  doc.rect(20, y - 4, 170, 20, "F");
  doc.setDrawColor(209, 143, 65);
  doc.setLineWidth(0.6);
  doc.rect(20, y - 4, 170, 20);
  doc.setFontSize(10);
  doc.setFont("helvetica", "bold");
  doc.setTextColor(133, 79, 11);
  doc.text("POSITIONS EN PORTEFEUILLE NON VÉRIFIÉES", 28, y + 3);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);
  doc.text("Liste des positions en portefeuille indisponible, contactez la Compliance avant toute opération.", 28, y + 10);
  return y + 24;
}

// Date du dernier dépôt de positions dans les PDF (attestation, vérification en masse)
function addPositionsFreshnessPdfLine(doc, y) {
  const info = positionsDepotInfo();
  if (!info) return y;
  const stale = info.ageDays > POSITIONS_MAX_AGE_DAYS;
  let text = info.author
    ? tf("pos_pdf_fresh_by", { date: formatDepotDate(info.date), author: info.author })
    : tf("pos_pdf_fresh", { date: formatDepotDate(info.date) });
  if (stale) text += ` ${tf("pos_pdf_fresh_old", { days: Math.floor(info.ageDays) })}`;
  if (y > 270) { doc.addPage(); y = 20; }
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(stale ? 154 : 110, stale ? 52 : 110, stale ? 18 : 110);
  const lines = doc.splitTextToSize(text, 170);
  doc.text(lines, 20, y);
  return y + lines.length * 4 + 4;
}

function updateRoleLabel() {
  const el = document.getElementById("nav-role-el");
  if (el) el.textContent = t(isCPL ? "nav_role_cpl" : "nav_role");
}

// Appelé par translations.js après chaque changement de langue
function onLanguageChanged() {
  updateRoleLabel();
  setPositionsStaticTexts();
  renderPositionsBanners();
  refreshPositionsPanelStatus();
  if (positionsPreview && !positionsDepositBusy) renderPositionsPreview(positionsPreview);
}

// ---- Interface de dépôt (profils Compliance uniquement) ----

const POSITIONS_ICON = `<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="7" width="20" height="14" rx="2" ry="2"/><path d="M16 21V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v16"/></svg>`;
const UPLOAD_ICON = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`;

// La carte et son panneau ne sont insérés dans le DOM que pour les profils Compliance
function insertPositionsDepositUI() {
  if (!isCPL || document.getElementById("positions-card")) return;
  const grid = document.querySelector("#tab-declaration .declare-grid");
  const giftForm = document.getElementById("gift-form-container");
  if (!grid || !giftForm) return;

  const card = document.createElement("div");
  card.className = "declare-card declare-card-clickable";
  card.id = "positions-card";
  card.innerHTML = `
    <div class="declare-card-icon" style="background:rgba(21,128,61,0.1);color:#15803d">${POSITIONS_ICON}</div>
    <span class="declare-card-badge badge badge-portefeuille">${PositionsCore.escapeHtml(PositionsCore.POSITION_TYPE)}</span>
    <h3 class="declare-card-title" id="pos-card-title-el"></h3>
    <p class="declare-card-desc" id="pos-card-desc-el"></p>
    <p class="declare-card-note" id="pos-card-note-el"></p>
    <div class="btn-form-open" style="margin-top:auto;background:#15803d">${UPLOAD_ICON}<span id="pos-card-btn-el"></span></div>`;
  card.addEventListener("click", togglePositionsForm);
  grid.appendChild(card);

  const panel = document.createElement("div");
  panel.id = "positions-form-container";
  panel.className = "hidden";
  panel.style.marginTop = "20px";
  panel.innerHTML = `
    <div class="form-card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px">
        <div>
          <h3 style="font-size:16px;font-weight:700;color:var(--navy)" id="pos-form-title-el"></h3>
          <p style="font-size:13px;color:var(--text-secondary);margin-top:3px" id="pos-form-subtitle-el"></p>
        </div>
        <button id="pos-form-close" style="background:none;border:none;cursor:pointer;color:var(--text-muted);font-size:22px;line-height:1">✕</button>
      </div>
      <div class="declare-alert" style="margin-bottom:14px">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        <span><span id="pos-form-format-el"></span> <a id="pos-template-link" href="templates/template-depot-positions.xlsx" download style="color:inherit;font-weight:600;text-decoration:underline;white-space:nowrap"></a></span>
      </div>
      <div id="pos-config-msg" class="form-message error hidden" style="margin-bottom:14px"></div>
      <div id="pos-dropzone" class="positions-dropzone">
        <p id="pos-drop-text-el"></p>
        <button type="button" id="pos-select-btn" class="btn-secondary">${UPLOAD_ICON}<span id="pos-select-btn-el"></span></button>
        <input type="file" id="pos-file-input" accept=".xlsx,.xls" multiple hidden />
      </div>
      <div id="pos-file-error" class="form-message error hidden" style="margin-top:14px"></div>
      <div id="pos-preview" class="hidden" style="margin-top:20px"></div>
      <div id="pos-progress" class="hidden" style="margin-top:20px">
        <p class="positions-busy-warning" id="pos-busy-el"></p>
        <div class="positions-progress-bar"><div class="positions-progress-fill" id="pos-progress-fill"></div></div>
        <p class="positions-progress-text" id="pos-progress-text"></p>
      </div>
      <div id="pos-result" class="hidden" style="margin-top:20px"></div>
    </div>`;
  giftForm.parentNode.insertBefore(panel, giftForm);

  document.getElementById("pos-form-close").addEventListener("click", togglePositionsForm);
  const input = document.getElementById("pos-file-input");
  document.getElementById("pos-select-btn").addEventListener("click", () => { if (!input.disabled) input.click(); });
  input.addEventListener("change", () => {
    const files = [...input.files];
    input.value = "";
    handlePositionsFiles(files);
  });
  const zone = document.getElementById("pos-dropzone");
  zone.addEventListener("dragover", e => { e.preventDefault(); if (!zone.classList.contains("disabled")) zone.classList.add("dragover"); });
  zone.addEventListener("dragleave", () => zone.classList.remove("dragover"));
  zone.addEventListener("drop", e => {
    e.preventDefault();
    zone.classList.remove("dragover");
    if (zone.classList.contains("disabled")) return;
    handlePositionsFiles([...(e.dataTransfer?.files || [])]);
  });

  setPositionsStaticTexts();
  refreshPositionsPanelStatus();
}

function setPositionsStaticTexts() {
  setText("pos-card-title-el", t("pos_card_title"));
  setText("pos-card-desc-el", t("pos_card_desc"));
  setText("pos-card-note-el", t("pos_card_note"));
  setText("pos-card-btn-el", t("pos_card_btn"));
  setText("pos-form-title-el", t("pos_form_title"));
  setText("pos-form-subtitle-el", t("pos_form_subtitle"));
  setText("pos-form-format-el", t("pos_form_format"));
  setText("pos-template-link", t("pos_template_link"));
  setText("pos-drop-text-el", t("pos_drop_text"));
  setText("pos-select-btn-el", t("pos_select_btn"));
  setText("pos-busy-el", t("pos_busy_warning"));
}

function togglePositionsForm() {
  const container = document.getElementById("positions-form-container");
  if (!container) return;
  const isHidden = container.classList.contains("hidden");
  if (isHidden) {
    container.classList.remove("hidden");
    container.scrollIntoView({ behavior: "smooth", block: "start" });
  } else if (!positionsDepositBusy) {
    container.classList.add("hidden");
  }
}

// Raison pour laquelle le dépôt est désactivé (null si le dépôt est possible)
function positionsDepositBlockedReason() {
  if (!siteIdHistory) return t("pos_err_site");
  if (positionsState.status === "loading") return t("pos_loading");
  if (!positionsTarget) return tf("pos_err_list", { detail: positionsTargetError || positionsState.reason || "" });
  if (!positionsColumnsOk) return tf("pos_err_columns", { detail: positionsColumnsError || "" });
  if (positionsState.status !== "ok") return tf("pos_err_read", { detail: positionsState.reason || "" });
  return null;
}

function refreshPositionsPanelStatus() {
  const msg = document.getElementById("pos-config-msg");
  if (!msg) return;
  const reason = positionsDepositBlockedReason();
  msg.textContent = reason || "";
  msg.classList.toggle("hidden", !reason);
  const disabled = !!reason || positionsDepositBusy;
  document.getElementById("pos-dropzone").classList.toggle("disabled", disabled);
  document.getElementById("pos-file-input").disabled = disabled;
  document.getElementById("pos-select-btn").disabled = disabled;
  const confirmBtn = document.getElementById("pos-confirm-btn");
  if (confirmBtn) confirmBtn.disabled = disabled || !positionsPreview || positionsPreview.isBlocked;
}

function showPositionsFileError(text) {
  const el = document.getElementById("pos-file-error");
  el.textContent = text;
  el.classList.remove("hidden");
}

async function handlePositionsFiles(files) {
  if (!isCPL || positionsDepositBusy || positionsDepositBlockedReason()) return;
  hideElement("pos-file-error");
  hideElement("pos-result");
  hideElement("pos-preview");
  positionsPreview = null;
  if (!files || files.length === 0) return;

  const bad = files.filter(f => !PositionsCore.ALLOWED_EXTENSIONS.includes(PositionsCore.fileExtension(f.name)));
  if (bad.length) { showPositionsFileError(tf("pos_err_format", { files: bad.map(f => f.name).join(", ") })); return; }
  if (typeof XLSX === "undefined") { showPositionsFileError(t("pos_err_sheetjs")); return; }

  const parsed = [];
  for (const f of files) {
    try {
      parsed.push(PositionsCore.parsePositionsWorkbook(XLSX, f.name, await f.arrayBuffer()));
    } catch (err) {
      console.error("Erreur lecture fichier positions:", err);
      showPositionsFileError(tf("pos_err_file", { file: f.name, detail: err.message }));
      return;
    }
  }
  const current = positionsState.rawItems.map(it => {
    const f = it.fields || {};
    return { nom: f.Title, isin: f.ISIN, strategie: f.Fonds, idDepot: f.IdDepot };
  });
  positionsPreview = PositionsCore.buildPreview(parsed, current);
  renderPositionsPreview(positionsPreview);
}

function renderPositionsPreview(p) {
  const esc = PositionsCore.escapeHtml;
  const el = document.getElementById("pos-preview");
  if (!el) return;
  const lineRef = r => tf("pos_line_ref", { file: esc(r.file), line: r.line });
  const cell = v => (v === "" ? `<span class="positions-empty">—</span>` : esc(v));
  const sumCounts = list => list.reduce((s, x) => s + x.count, 0);
  const detailsList = (title, items) => items.length
    ? `<details class="positions-details"><summary>${title}</summary><ul>${items.map(i => `<li>${i}</li>`).join("")}</ul></details>`
    : "";
  const posTable = list => `
    <div class="table-wrapper positions-scroll"><table class="data-table">
      <thead><tr><th>${t("pos_col_name")}</th><th>${t("pos_col_isin")}</th><th>${t("pos_col_strategy")}</th><th>${t("pos_col_count")}</th></tr></thead>
      <tbody>${list.map(x => `<tr><td>${cell(x.nom)}</td><td style="font-family:monospace">${cell(x.isin)}</td><td>${cell(x.strategie)}</td><td>${x.count}</td></tr>`).join("")}</tbody>
    </table></div>`;

  let html = `<h4 class="positions-section-title">${t("pos_prev_title")}</h4>
    <div class="positions-summary">
      <span class="positions-chip"><strong>${tf("pos_prev_total", { n: p.total })}</strong></span>
      <span class="positions-chip">${tf("pos_prev_current", { n: p.currentTotal })}</span>
    </div>
    <div class="positions-grid-2">
      <div>
        <h5 class="positions-subtitle">${t("pos_prev_by_file")}</h5>
        <div class="table-wrapper"><table class="data-table">
          <thead><tr><th>${t("pos_col_file")}</th><th>${t("pos_col_rows")}</th></tr></thead>
          <tbody>${p.byFile.map(f => `<tr><td>${esc(f.fileName)}</td><td>${f.count}</td></tr>`).join("")}</tbody>
        </table></div>
      </div>
      <div>
        <h5 class="positions-subtitle">${t("pos_prev_by_strategy")}</h5>
        <div class="table-wrapper positions-scroll"><table class="data-table">
          <thead><tr><th>${t("pos_col_strategy")}</th><th>${t("pos_col_rows")}</th></tr></thead>
          <tbody>${p.byStrategy.map(s => `<tr><td>${s.strategie === "" ? `<em>${t("pos_prev_no_strategy")}</em>` : esc(s.strategie)}</td><td>${s.count}</td></tr>`).join("")}</tbody>
        </table></div>
      </div>
    </div>`;

  if (p.isBlocked) {
    const fieldLabel = f => t(`pos_field_${f}`);
    html += `<div class="positions-box positions-box-error">
      <strong>${t("pos_block_title")}</strong>
      ${p.blocking.empty ? `<p>${t("pos_block_empty")}</p>` : ""}
      ${p.blocking.tooLong.length ? `<p>${tf("pos_block_toolong", { n: p.blocking.tooLong.length, max: PositionsCore.MAX_CELL_LENGTH })}</p>
        <ul>${p.blocking.tooLong.map(x => `<li>${lineRef(x.row)} : ${fieldLabel(x.field)} (${x.length})</li>`).join("")}</ul>` : ""}
    </div>`;
  }

  if (p.manualRows.length) {
    html += `<div class="positions-box positions-box-error">
      <strong>${tf("pos_manual_title", { n: p.manualRows.length })}</strong>
      <p>${t("pos_manual_text")}</p>
      ${detailsList(tf("pos_manual_list", { n: p.manualRows.length }), p.manualRows.map(r => `${cell(r.nom)} / ${cell(r.isin)} / ${cell(r.strategie)}`))}
    </div>`;
  }

  const w = p.warnings;
  const hasWarnings = w.isinSuspect.length || w.duplicates.length || w.emptyName.length || w.emptyStrategy.length;
  if (hasWarnings) {
    html += `<div class="positions-box positions-box-warning">
      <strong>${t("pos_warn_title")}</strong>
      ${detailsList(tf("pos_warn_isin", { n: w.isinSuspect.length }), w.isinSuspect.map(r => `${lineRef(r)} : ${esc(r.isin)}`))}
      ${detailsList(tf("pos_warn_dup", { n: w.duplicates.length }), w.duplicates.map(g => `${cell(g[0].nom)} / ${cell(g[0].isin)} / ${cell(g[0].strategie)} : ${g.map(lineRef).join(" ; ")}`))}
      ${detailsList(tf("pos_warn_noname", { n: w.emptyName.length }), w.emptyName.map(lineRef))}
      ${detailsList(tf("pos_warn_nostrategy", { n: w.emptyStrategy.length }), w.emptyStrategy.map(lineRef))}
    </div>`;
  }

  html += `<h5 class="positions-subtitle">${t("pos_diff_title")}</h5>
    <details class="positions-details"><summary>${tf("pos_diff_added", { n: sumCounts(p.added) })}</summary>${p.added.length ? posTable(p.added) : ""}</details>
    <details class="positions-details"><summary>${tf("pos_diff_removed", { n: sumCounts(p.removed) })}</summary>${p.removed.length ? posTable(p.removed) : ""}</details>

    <h5 class="positions-subtitle">${t("pos_prev_rows")}</h5>
    <div class="table-wrapper positions-scroll"><table class="data-table">
      <thead><tr><th>${t("pos_col_name")}</th><th>${t("pos_col_isin")}</th><th>${t("pos_col_strategy")}</th><th>${t("pos_col_line")}</th></tr></thead>
      <tbody>${p.rows.map(r => `<tr><td>${cell(r.nom)}</td><td style="font-family:monospace">${cell(r.isin)}</td><td>${cell(r.strategie)}</td><td>${lineRef(r)}</td></tr>`).join("")}</tbody>
    </table></div>

    <div class="positions-box positions-box-info">${tf("pos_confirm_reminder", { n: p.total })}</div>
    <div class="form-actions" style="justify-content:flex-end;gap:10px">
      <button type="button" id="pos-cancel-btn" class="btn-secondary">${t("pos_cancel_btn")}</button>
      <button type="button" id="pos-confirm-btn" class="btn-primary" style="background:#15803d">${t("pos_confirm_btn")}</button>
    </div>`;

  el.innerHTML = html;
  el.classList.remove("hidden");
  document.getElementById("pos-cancel-btn").addEventListener("click", () => {
    positionsPreview = null;
    el.classList.add("hidden");
    el.innerHTML = "";
  });
  document.getElementById("pos-confirm-btn").addEventListener("click", runPositionsDeposit);
  refreshPositionsPanelStatus();
}

// Adaptateur Graph injecté dans PositionsCore.runReplacement
const positionsGraphAdapter = {
  async postBatch(body) {
    // Troisième verrou : chaque sous-requête est revérifiée juste avant l'envoi
    (body.requests || []).forEach(r => PositionsCore.assertBatchRequestAllowed(r, positionsTarget));
    const token = await acquireGraphTokenNoRedirect();
    const response = await fetch("https://graph.microsoft.com/v1.0/$batch", {
      method: "POST",
      headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const json = await response.json().catch(() => null);
    return { status: response.status, retryAfter: response.headers.get("Retry-After"), json };
  },
  listItems(target) { return listPositionItems(target); }
};

function positionsBeforeUnload(e) {
  e.preventDefault();
  e.returnValue = "";
  return "";
}

function updatePositionsProgress(p) {
  const key = p.phase === "delete" ? "pos_progress_delete" : p.phase === "rollback" ? "pos_progress_rollback" : "pos_progress_create";
  const pct = p.total ? Math.round((p.done / p.total) * 100) : 100;
  const fill = document.getElementById("pos-progress-fill");
  if (fill) fill.style.width = `${pct}%`;
  setText("pos-progress-text", tf(key, { done: p.done, total: p.total }));
}

async function runPositionsDeposit() {
  if (!isCPL || positionsDepositBusy) return;
  const preview = positionsPreview;
  if (!preview || preview.isBlocked) return;
  const blocked = positionsDepositBlockedReason();
  if (blocked) { refreshPositionsPanelStatus(); return; }

  positionsDepositBusy = true;
  refreshPositionsPanelStatus();
  document.getElementById("pos-cancel-btn").disabled = true;
  window.addEventListener("beforeunload", positionsBeforeUnload);
  hideElement("pos-result");
  showElement("pos-progress");
  updatePositionsProgress({ phase: "create", done: 0, total: preview.rows.length });

  let result = null;
  let fatal = null;
  let lock = null;
  let historyNote = null;
  lastBackupFile = null;
  try {
    const idDepot = PositionsCore.generateIdDepot();
    lock = await acquirePositionsLock(idDepot, preview, preview.byFile.map(f => f.fileName));
    result = await PositionsCore.runReplacement({
      graph: positionsGraphAdapter,
      target: positionsTarget,
      rows: preview.rows,
      idDepot,
      onProgress: updatePositionsProgress,
      beforeDelete: oldItems => backupPositionsBeforeDelete(idDepot, oldItems)
    });
  } catch (err) {
    if (!(err instanceof PositionsLockError)) console.error("Erreur dépôt positions:", err);
    fatal = err;
  } finally {
    if (!(fatal instanceof PositionsLockError)) historyNote = await finalizePositionsHistory(lock, result, fatal, preview);
    window.removeEventListener("beforeunload", positionsBeforeUnload);
    positionsDepositBusy = false;
  }

  hideElement("pos-progress");
  positionsPreview = null;
  const previewEl = document.getElementById("pos-preview");
  previewEl.classList.add("hidden");
  previewEl.innerHTML = "";
  renderPositionsResult(result, fatal, historyNote);
  await loadRestrictedList();
}

function renderPositionsResult(result, fatal, historyNote) {
  const el = document.getElementById("pos-result");
  el.textContent = "";
  const box = document.createElement("div");
  const lines = [];
  if (fatal && fatal.name === "PositionsLockError") {
    box.className = "positions-box positions-box-warning";
    lines.push(fatal.message);
    lines.push(t("pos_lock_nothing_changed"));
  } else if (fatal) {
    box.className = "positions-box positions-box-error";
    lines.push(tf("pos_result_fatal", { detail: fatal.message }));
  } else if (result.ok && result.deleteFailed === 0) {
    box.className = "positions-box positions-box-success";
    lines.push(tf("pos_result_ok", { created: result.created, deleted: result.deleted }));
  } else if (result.ok) {
    box.className = "positions-box positions-box-warning";
    lines.push(tf("pos_result_partial_delete", { created: result.created, deleted: result.deleted, failed: result.deleteFailed }));
  } else {
    box.className = "positions-box positions-box-error";
    lines.push(tf("pos_result_failed", { detail: result.failure.message }));
    lines.push(tf("pos_result_rollback", { rolledBack: result.rolledBack }));
    if (result.rollbackFailed > 0 || result.rollbackReadFailed) {
      lines.push(tf("pos_result_rollback_incomplete", { failed: result.rollbackFailed }));
    }
  }
  if (result && result.ok && lastBackupFile) lines.push(tf("pos_result_backup", { file: lastBackupFile }));
  if (historyNote) lines.push(historyNote);
  lines.forEach(text => {
    const p = document.createElement("p");
    p.textContent = text;
    box.appendChild(p);
  });
  el.appendChild(box);
  el.classList.remove("hidden");
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
