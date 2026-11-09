# Compliance Portal — Eiffel Investment Group

Portail statique (Azure Static Web App) avec des Azure Functions (`api/`) qui font le pont vers
Microsoft Graph / SharePoint et vers des API tierces, en gardant les secrets côté serveur.

## Configuration — variables d'environnement des Azure Functions

Ces variables se définissent dans les **Application Settings** de l'Azure Static Web App / Function
App, depuis le Portail Azure. Elles ne doivent **jamais** être commitées dans le dépôt (ni dans le
front-end, ni dans `api/local.settings.json`).

| Variable | Utilisée par | Rôle |
|---|---|---|
| `OPENAI_API_KEY` | Functions Doc Review | Clé de l'API OpenAI (Doc Review). |
| `OPENAI_BASE_URL` | Functions Doc Review | *(Optionnel)* URL de base de l'API OpenAI. Par défaut `https://api.openai.com/v1` ; `https://eu.api.openai.com/v1` pour le point d'accès UE (projet OpenAI configuré en résidence UE). |
| `OPENAI_MODEL_PAGES` | `api/docreview-start` | *(Optionnel)* Modèle de l'étape 1 (analyse par lots de pages). Par défaut `gpt-5.6-terra`. |
| `OPENAI_MODEL_SYNTHESE` | `api/docreview-start` | *(Optionnel)* Modèle de l'étape 2 (synthèse). Par défaut `gpt-5.6-terra`. |
| `OPENAI_EFFORT_PAGES` | `api/docreview-start` | *(Optionnel)* Effort de raisonnement de l'étape 1 (`none`, `minimal`, `low`, `medium`, `high`, `xhigh`, `max` selon le modèle). Par défaut `low`. |
| `OPENAI_EFFORT_SYNTHESE` | `api/docreview-start` | *(Optionnel)* Effort de raisonnement de l'étape 2. Par défaut `medium`. |

Le navigateur ne choisit jamais le modèle : seuls les modèles configurés ci-dessus sont utilisés.

### Authentification des Functions

Toutes les Functions exigent le jeton Microsoft Graph de l'utilisateur connecté (en-tête
`X-EIG-Graph-Token` ou `Authorization: Bearer`). La Function appelle `/me` avec ce jeton et refuse
la requête (401) si l'appel échoue ou si le `userPrincipalName` ne se termine pas par
`@eiffel-ig.com`. Le résultat est mis en cache 5 minutes par jeton (code : `api/shared/auth.js`).

### Développement local

Pour tester une Function en local (Azure Functions Core Tools / SWA CLI), créer un fichier
`api/local.settings.json` (déjà couvert par `.gitignore` — **ne pas le committer**) :

```json
{
  "IsEncrypted": false,
  "Values": {
    "FUNCTIONS_WORKER_RUNTIME": "node",
    "OPENAI_API_KEY": "..."
  }
}
```

## Doc Review : relecture IA de la documentation marketing

- **Formulaire** : type de fonds obligatoire (liste modifiable dans `docreview-options.js`, qui
  indique aussi les fonds réservés aux professionnels pour l'avertissement non bloquant en Retail)
  et public visé (Retail ou Professionnel). La nature du document est déterminée par l'IA.
- **Lecture complète** : PDF, PPTX, DOCX (un email ou un post s'analyse enregistré en PDF ou DOCX). Les PPTX et DOCX sont convertis en PDF par
  SharePoint (`?format=pdf`) ; chaque page est rendue en image (pdf.js) et son texte extrait. Toutes
  les pages sont analysées. Les slides masquées et les pages sans couche texte sont signalées.
- **Analyse en deux étapes**, en mode background de l'API Responses d'OpenAI (les Functions intégrées
  coupent les requêtes après 45 secondes) : `api/docreview-start` démarre une étape,
  `api/docreview-status` renvoie son état, `api/docreview-cancel` l'annule. Étape 1 : relevé fidèle
  par lots de pages (images + texte). Étape 2 : analyse d'ensemble. Sorties structurées en JSON strict (schémas dans `api/shared/docreview-logic.js`).
- **Prompts** : `api/shared/prompts-docreview.md`, fichier lisible et commenté, lu à chaque analyse
  (prompt de synthèse, prompt par lots de pages, annexe des règles internes Eiffel). Les prompts
  sont construits côté serveur ; le navigateur n'envoie que les données du document.
- **Résultat** : compréhension du document, appréciation globale, points forts, remarques
  hiérarchisées (À traiter, Recommandé, Suggestion) avec niveau de confiance, points à vérifier.
  Aucun score calculé : la colonne `Score` vaut 3 (Prêt à soumettre), 2 (Quelques ajustements
  conseillés) ou 1 (À retravailler) ; `NbBloquants` compte les remarques "À traiter" et
  `NbAttention` les remarques "Recommandé".
- **Traçabilité** : liste "Doc Review Analyses" (dont `TypeFonds` : type de fonds choisi,
  `TypeDocument` : nature détectée par l'IA) et bibliothèque "Doc Review" (site CPLDashboard), un
  dossier par analyse (`DR-AAAA-NNNN`) contenant le fichier d'origine, le PDF converti, le rapport PDF
  et le rapport JSON. Statuts : En cours, Analysé, Erreur (motif dans le rapport JSON), Soumis.

## Restricted List : positions en portefeuille

La Restricted List (`index.html`) réunit trois sources : "NDA List" et "Information privilégiée"
(site `/sites/CPL`) et la liste **"positions portefeuilles"** (site `/sites/CPLDashboard`, colonnes
`Title`, `ISIN`, `Fonds` (nom d'affichage "Stratégie"), `IdDepot`). La stratégie des positions
s'affiche dans la colonne Équipe du tableau.

- **Dépôt** : onglet *Déclarer*, carte *En portefeuille*, visible des seuls profils Compliance
  (liste `CPL_EMAILS` dans `cpl-config.js`, contrôle d'affichage uniquement : la vraie protection
  est la permission SharePoint de la liste). Formats acceptés : `.xlsx` et `.xls` (première
  feuille, ligne 1 = en-têtes, colonne A = nom du titre, B = ISIN, C = stratégie). Les fichiers
  déposés ensemble remplacent la totalité des positions, après un aperçu obligatoire.
- **Remplacement sécurisé** : création de toutes les lignes avec un nouvel `IdDepot`, vérification,
  puis seulement suppression des anciennes lignes. En cas d'échec, les lignes partielles sont
  supprimées et l'ancienne liste est conservée. Toutes les suppressions passent par une garde
  unique (`PositionsCore.assertDeletable`) qui refuse toute autre liste.
- **Indisponibilité** : si la liste ne peut pas être lue, NDA et informations privilégiées restent
  affichées, avec un bandeau invitant à contacter la Compliance (également repris dans
  l'attestation PDF et la vérification en masse).
- Code : logique testable dans `positions-core.js` (sans dépendance à Graph ni au DOM),
  intégration dans `app.js`. Lecture Excel : SheetJS 0.20.3 (`cdn.sheetjs.com`, empreinte SRI).

## CounterFlow : circuit de validation KYC

- **Proposer** : tout collaborateur peut créer une contrepartie ou demander la mise à jour de son
  KYC (niveau de risque, lien, date d'expiration, pièces). La proposition est stockée dans les
  colonnes `*_Pending` ; le dossier officiel n'est modifié qu'à la validation. Une seule demande
  à la fois par contrepartie ; les noms en double sont refusés à la création.
- **Décider** (profils Compliance) : Valider, Demander un complément ou Rejeter (message
  obligatoire). L'auteur d'une demande ne peut pas la valider lui-même. Après un complément ou un
  rejet, la proposition est conservée et le demandeur la corrige via « Corriger et resoumettre ».
- **Notifications** : email à `cpl@eiffel-ig.com` à chaque soumission, email au demandeur à chaque
  décision (envoyés depuis la boîte de l'utilisateur, non bloquants).
- **Statuts** : deux informations distinctes, la validité du dossier officiel (à jour, bientôt
  expiré, expiré, non validé) et l'état de la demande (en attente, complément demandé, rejetée).
  Les relances portent sur la validité, demande en cours ou non ; les contreparties archivées et
  les créations jamais validées en sont exclues.
- **Historique** : chaque action est ajoutée à la colonne `HistoriqueKYC` (JSON, ajout seul) et
  affichée sur la fiche.
- **Colonnes de la liste Brokers ajoutées pour ce circuit** (détectées au chargement ; tant
  qu'une colonne manque, l'ancien fonctionnement correspondant s'applique) :

| Colonne (nom interne) | Type SharePoint | Rôle |
|---|---|---|
| `NiveauRisque_Pending` | Une seule ligne de texte | Niveau de risque proposé, soumis à validation |
| `HistoriqueKYC` | Plusieurs lignes de texte (texte brut) | Historique des actions |
| `Archive` | Oui/Non | Contrepartie archivée (masquée, non relancée) |
| `StatutValidationKYC` | Choix existant | Ajouter la valeur `Complément demandé` |

Le contrôle du profil Compliance est fait dans le navigateur : la protection réelle repose sur les
permissions SharePoint de la liste Brokers.
