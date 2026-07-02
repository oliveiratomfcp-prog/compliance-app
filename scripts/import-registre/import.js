#!/usr/bin/env node
/**
 * Import ponctuel du Registre déontologique Excel vers les listes SharePoint
 * de DeontoHub (site CPLDashboard).
 *
 * ─────────────────────────────────────────────────────────────────────────
 * IMPORTANT — À LIRE AVANT D'EXÉCUTER
 * ─────────────────────────────────────────────────────────────────────────
 * Le fichier source a des en-têtes fusionnés et irréguliers sur plusieurs
 * lignes (confirmé en inspectant l'aperçu du fichier réel). Ce script ne
 * devine donc PAS le mapping des colonnes : il fournit un mode --inspect
 * pour que tu voies la structure réelle, puis tu ajustes toi-même la
 * constante COLUMN_MAP / OBLIGATION_COLUMNS ci-dessous en conséquence
 * avant de lancer --dry-run puis l'import réel. Les valeurs pré-remplies
 * ci-dessous sont des hypothèses de départ, PAS des valeurs vérifiées.
 *
 * Workflow recommandé :
 *   1. npm install
 *   2. npm run inspect   -> lis la sortie, identifie les vraies colonnes
 *   3. Ajuste COLUMN_MAP et OBLIGATION_COLUMNS plus bas dans ce fichier
 *   4. npm run dry-run    -> vérifie les objets JSON qui seraient envoyés
 *   5. npm run import     -> écrit réellement dans SharePoint
 *
 * Authentification : device code flow MSAL (mêmes clientId/tenant que le
 * reste de l'app, déjà consentis pour Sites.ReadWrite.All). Aucune valeur
 * secrète nécessaire : le script affiche un code à saisir sur
 * https://microsoft.com/devicelogin, avec le compte Compliance habituel.
 * Ne commite jamais de fichier .env ni de données réelles issues d'un
 * run de ce script.
 */

const fs = require('fs');
const XLSX = require('xlsx');
const fetch = require('node-fetch');
const { PublicClientApplication } = require('@azure/msal-node');

// ─────────────────────────────────────────────────────────────────────────
// CONFIGURATION (variables d'environnement — jamais d'identifiants en dur)
// ─────────────────────────────────────────────────────────────────────────
const TENANT_ID  = process.env.TENANT_ID  || 'eb5288b6-5ccb-4f5d-88fb-346db1317a4f';
const CLIENT_ID  = process.env.CLIENT_ID  || '56c83db2-7fc6-4826-8073-82077e91ae27'; // même app publique que deontohub.html
const EXCEL_PATH = process.env.EXCEL_PATH; // chemin local vers une copie téléchargée du xlsx

const SP_HOSTNAME   = process.env.SP_HOSTNAME   || 'eiffelinvestmentgroup.sharepoint.com';
const SP_SITE_PATH  = process.env.SP_SITE_PATH  || '/sites/CPLDashboard';
const LIST_COLLAB      = process.env.LIST_COLLAB      || 'Registre Collaborateurs';
const LIST_OBLIGATIONS = process.env.LIST_OBLIGATIONS || 'Registre Obligations';

const SHEETS = {
  actif      : 'Registre (actif)',
  depart     : 'Départ',
  stagiaires : 'Stagiaires'
};

// ─────────────────────────────────────────────────────────────────────────
// MAPPING DES COLONNES — À AJUSTER APRÈS `npm run inspect`
// Index de colonne 0-based (colonne A = 0, B = 1, ...).
// La ligne d'en-tête réelle (headerRow) est aussi à ajuster : c'est le
// numéro de ligne (0-based) où commencent les vraies données, juste après
// le bloc d'en-têtes fusionnés.
// ─────────────────────────────────────────────────────────────────────────
const COLUMN_MAP = {
  actif: {
    headerRow: 2,           // ⚠ hypothèse — à confirmer avec --inspect
    nom: 0,                 // Prénom + Nom (ou 2 colonnes séparées à concaténer, voir parseRow)
    prenom: null,           // si Prénom/Nom sont deux colonnes distinctes, renseigner ici et mettre nom=colonne "Nom"
    email: null,            // pas vu de colonne email dans l'aperçu — à confirmer, sinon laisser null
    departement: 1,
    fonction: 2,            // fonction clé "principale" en texte libre dans le fichier source
    dateEntree: 3,
    typeContrat: null,
    statutDeclareAMF: null,
    certificationRequise: null
  },
  depart: {
    headerRow: 1,
    nom: 0,
    departement: 1,
    dateEntree: 2,
    dateSortie: 3
  },
  stagiaires: {
    headerRow: 1,
    nom: 0,
    departement: 1,
    dateEntree: 2
  }
};

// Colonnes "obligation" en format large (une colonne par catégorie/année)
// à convertir en lignes Registre Obligations. Chaque entrée : la colonne
// contenant soit une date, soit un texte/statut pour cette catégorie.
// ⚠ Liste indicative à partir du texte du brief — à vérifier et compléter
// avec les vrais index de colonnes vus dans --inspect.
const OBLIGATION_COLUMNS = [
  // { columnIndex: 4, categorie: 'Formation AML' },
  // { columnIndex: 5, categorie: 'Formation MAR' },
  // { columnIndex: 6, categorie: 'Certification AMF' },
  // { columnIndex: 7, categorie: 'Certification Finance Durable' },
  // { columnIndex: 8, categorie: 'Mandat Externe' },
];

// ─────────────────────────────────────────────────────────────────────────
// CLI
// ─────────────────────────────────────────────────────────────────────────
const MODE = process.argv.includes('--inspect') ? 'inspect'
  : process.argv.includes('--dry-run') ? 'dry-run'
  : 'import';

function fail(msg) { console.error(`\n✕ ${msg}\n`); process.exit(1); }

if (!EXCEL_PATH) fail('Variable d\'environnement EXCEL_PATH manquante (chemin local vers le fichier .xlsx).');
if (!fs.existsSync(EXCEL_PATH)) fail(`Fichier introuvable : ${EXCEL_PATH}`);

// ─────────────────────────────────────────────────────────────────────────
// INSPECTION — dump brut des premières lignes de chaque feuille
// ─────────────────────────────────────────────────────────────────────────
function inspect(workbook) {
  console.log(`\nFeuilles trouvées dans le classeur : ${workbook.SheetNames.join(', ')}\n`);
  Object.entries(SHEETS).forEach(([key, sheetName]) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) { console.log(`⚠ Feuille "${sheetName}" (${key}) introuvable dans ce classeur.\n`); return; }
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    console.log(`── ${sheetName} (${key}) — ${rows.length} lignes ──`);
    rows.slice(0, 8).forEach((row, i) => {
      console.log(`  [${i}] ${row.map((c, ci) => `(${ci})${String(c).slice(0,28)}`).join(' | ')}`);
    });
    console.log('');
  });
  console.log('Ajuste COLUMN_MAP et OBLIGATION_COLUMNS dans import.js à partir de ces index de colonnes,\npuis relance avec --dry-run.\n');
}

// ─────────────────────────────────────────────────────────────────────────
// PARSING — feuille -> { collaborateurs, obligations }
// ─────────────────────────────────────────────────────────────────────────
function parseSheet(workbook, sheetKey, statutParDefaut) {
  const sheetName = SHEETS[sheetKey];
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) { console.warn(`⚠ Feuille "${sheetName}" introuvable, ignorée.`); return { collaborateurs: [], obligations: [] }; }
  const map = COLUMN_MAP[sheetKey];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  const dataRows = rows.slice(map.headerRow + 1).filter(r => r[map.nom] && String(r[map.nom]).trim());

  const collaborateurs = [];
  const obligations = [];

  dataRows.forEach(row => {
    const nom = map.prenom != null
      ? `${row[map.prenom] || ''} ${row[map.nom] || ''}`.trim()
      : String(row[map.nom]).trim();

    const collab = {
      nom,
      email: map.email != null ? String(row[map.email] || '').trim() : '',
      departement: map.departement != null ? String(row[map.departement] || '').trim() : '',
      dateEntree: map.dateEntree != null ? row[map.dateEntree] : '',
      dateSortie: map.dateSortie != null ? row[map.dateSortie] : '',
      typeContrat: map.typeContrat != null ? String(row[map.typeContrat] || '').trim() : '',
      statut: statutParDefaut,
      fonctionCle: map.fonction != null && row[map.fonction] ? [String(row[map.fonction]).trim()] : [],
      statutDeclareAMF: map.statutDeclareAMF != null ? String(row[map.statutDeclareAMF] || '').trim() : '',
      certificationRequise: map.certificationRequise != null ? !!row[map.certificationRequise] : false
    };
    collaborateurs.push(collab);

    OBLIGATION_COLUMNS.forEach(oc => {
      const val = row[oc.columnIndex];
      if (!val || !String(val).trim()) return;
      obligations.push({
        collaborateurNom: nom,
        categorie: oc.categorie,
        annee: new Date().getFullYear(),
        dateEvenement: val,
        statut: 'OK',
        detail: '',
        commentaire: `Importé automatiquement depuis "${sheetName}"`
      });
    });
  });

  return { collaborateurs, obligations };
}

// ─────────────────────────────────────────────────────────────────────────
// AUTH — device code flow, aucun secret
// ─────────────────────────────────────────────────────────────────────────
async function getAccessToken() {
  const pca = new PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` }
  });
  const result = await pca.acquireTokenByDeviceCode({
    scopes: ['Sites.ReadWrite.All'],
    deviceCodeCallback: (resp) => console.log(`\n${resp.message}\n`)
  });
  return result.accessToken;
}

async function graphGet(token, endpoint) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`GET ${endpoint} -> ${res.status} ${await res.text()}`);
  return res.json();
}
async function graphPost(token, endpoint, body) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`POST ${endpoint} -> ${res.status} ${await res.text()}`);
  return res.json();
}

// ─────────────────────────────────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────────────────────────────────
async function main() {
  const workbook = XLSX.readFile(EXCEL_PATH);

  if (MODE === 'inspect') { inspect(workbook); return; }

  const actif      = parseSheet(workbook, 'actif', 'Actif');
  const depart     = parseSheet(workbook, 'depart', 'Parti');
  const stagiaires = parseSheet(workbook, 'stagiaires', 'Stagiaire');

  const allCollaborateurs = [...actif.collaborateurs, ...depart.collaborateurs, ...stagiaires.collaborateurs];
  const allObligations    = [...actif.obligations, ...depart.obligations, ...stagiaires.obligations];

  console.log(`Collaborateurs détectés : ${allCollaborateurs.length} (Actif: ${actif.collaborateurs.length}, Parti: ${depart.collaborateurs.length}, Stagiaire: ${stagiaires.collaborateurs.length})`);
  console.log(`Obligations détectées   : ${allObligations.length}`);

  if (MODE === 'dry-run') {
    console.log('\n--- Aperçu des 5 premiers collaborateurs ---');
    console.log(JSON.stringify(allCollaborateurs.slice(0, 5), null, 2));
    console.log('\n--- Aperçu des 5 premières obligations ---');
    console.log(JSON.stringify(allObligations.slice(0, 5), null, 2));
    console.log('\nAucune écriture effectuée (--dry-run). Relance sans option pour importer réellement.');
    return;
  }

  if (!allCollaborateurs.length) fail('Aucun collaborateur détecté : vérifie COLUMN_MAP avant de relancer.');

  console.log('\nConnexion à Microsoft Graph (device code)...');
  const token = await getAccessToken();

  const site = await graphGet(token, `/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`);
  const collabLists = await graphGet(token, `/sites/${site.id}/lists?$filter=displayName eq '${encodeURIComponent(LIST_COLLAB)}'`);
  const obligLists  = await graphGet(token, `/sites/${site.id}/lists?$filter=displayName eq '${encodeURIComponent(LIST_OBLIGATIONS)}'`);
  const listCollabId = collabLists.value[0]?.id;
  const listObligId  = obligLists.value[0]?.id;
  if (!listCollabId) fail(`Liste "${LIST_COLLAB}" introuvable sur ${SP_SITE_PATH}.`);
  if (!listObligId)  fail(`Liste "${LIST_OBLIGATIONS}" introuvable sur ${SP_SITE_PATH}.`);

  const nomToItemId = {};
  const errors = [];

  console.log(`\nImport de ${allCollaborateurs.length} collaborateur(s)...`);
  for (const c of allCollaborateurs) {
    try {
      const created = await graphPost(token, `/sites/${site.id}/lists/${listCollabId}/items`, {
        fields: {
          Title: c.nom,
          Email: c.email,
          Departement: c.departement,
          DateEntree: c.dateEntree ? new Date(c.dateEntree).toISOString() : null,
          DateSortie: c.dateSortie ? new Date(c.dateSortie).toISOString() : null,
          TypeContrat: c.typeContrat,
          Statut: c.statut,
          FonctionCle: c.fonctionCle,
          StatutDeclareAMF: c.statutDeclareAMF,
          CertificationRequise: c.certificationRequise
        }
      });
      nomToItemId[c.nom] = created.id;
      process.stdout.write('.');
    } catch(e) {
      errors.push(`Collaborateur "${c.nom}": ${e.message}`);
      process.stdout.write('x');
    }
  }

  console.log(`\n\nImport de ${allObligations.length} obligation(s)...`);
  for (const o of allObligations) {
    const collabId = nomToItemId[o.collaborateurNom];
    if (!collabId) { errors.push(`Obligation "${o.categorie}" ignorée : collaborateur "${o.collaborateurNom}" non trouvé.`); continue; }
    try {
      await graphPost(token, `/sites/${site.id}/lists/${listObligId}/items`, {
        fields: {
          Title: `${o.collaborateurNom} — ${o.categorie}`,
          Collaborateur: collabId,
          Categorie: o.categorie,
          Annee: o.annee,
          DateEvenement: o.dateEvenement ? new Date(o.dateEvenement).toISOString() : null,
          Statut: o.statut,
          Detail: o.detail,
          Commentaire: o.commentaire
        }
      });
      process.stdout.write('.');
    } catch(e) {
      errors.push(`Obligation "${o.collaborateurNom} — ${o.categorie}": ${e.message}`);
      process.stdout.write('x');
    }
  }

  console.log(`\n\nTerminé. ${Object.keys(nomToItemId).length}/${allCollaborateurs.length} collaborateur(s) importé(s).`);
  if (errors.length) {
    console.log(`\n${errors.length} erreur(s) :`);
    errors.forEach(e => console.log(`  - ${e}`));
    process.exitCode = 1;
  }
}

main().catch(e => fail(e.stack || e.message));
