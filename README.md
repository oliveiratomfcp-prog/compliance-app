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
- **Modèle Excel** : lien « Télécharger le modèle Excel » dans le formulaire de dépôt
  (`templates/template-depot-positions.xlsx`, en-têtes seuls : Nom du titre, ISIN, Stratégie). Les
  fichiers du site étant accessibles sans connexion, ce dossier ne doit jamais contenir de vraies positions.
- **Verrou et historique des dépôts** : chaque dépôt crée une entrée « En cours » dans la liste
  « Historique positions » (site CPLDashboard), relue juste après pour départager deux dépôts
  simultanés (le plus ancien gagne, l'autre est refusé). Un verrou « En cours » de plus d'une heure
  est considéré abandonné. L'entrée est clôturée avec le résultat (Terminé, Terminé avec erreurs,
  Échec, Annulé), l'auteur, les fichiers et les nombres de lignes.
- **Sauvegarde avant suppression** : avant de supprimer l'ancienne liste, elle est enregistrée en CSV
  dans la bibliothèque « Sauvegardes positions ». Si l'enregistrement échoue, rien n'est supprimé et
  les nouvelles lignes sont retirées. Restauration manuelle : rouvrir le CSV et le redéposer.
- **Lignes ajoutées à la main** : les lignes sans `IdDepot` seraient supprimées par un remplacement ;
  l'aperçu les liste dans un encadré rouge avant confirmation (elles figurent dans la sauvegarde).
- **Fraîcheur** : la date (et l'auteur, si l'historique existe) du dernier dépôt s'affiche au-dessus du
  tableau et sur les PDF (attestation, vérification en masse). Au-delà de `POSITIONS_MAX_AGE_DAYS`
  (7 jours, `app.js`) ou si la liste est vide, un bandeau d'alerte s'affiche pour tous.
- **Rafraîchissement automatique** : la Restricted List se recharge toutes les 2 minutes (onglet visible,
  hors dépôt en cours). La recherche et le tri sont conservés ; en cas d'échec, les données affichées
  sont conservées.
- Colonnes de la liste « Historique positions » (seule `Statut` est indispensable ; les autres sont
  renseignées si elles existent) : `Title` (IdDepot), `Statut`, `Auteur`, `Fichiers`, `NbLignes`,
  `NbAnciennes`, `NbAjoutees`, `NbRetirees`, `NbManuelles`, `Sauvegarde`, `Detail` (texte ; les colonnes
  `Nb...` peuvent être de type Nombre). Si la liste ou la bibliothèque manque, le dépôt reste possible
  mais sans verrou ni sauvegarde, et la Compliance en est avertie.
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

## Tracker CPL : recherche IA et demandes similaires

Deux fonctions s'appuient sur l'API OpenAI déjà configurée pour Doc Review (mêmes `OPENAI_API_KEY` et
`OPENAI_BASE_URL`), via les Azure Functions `api/ai-search` et `api/ai-similar` (mêmes règles
d'authentification que les autres Functions, 15 appels par minute et par utilisateur).

- **Recherche IA** : bouton « ✦ Recherche IA » (ou Entrée) à côté du champ « Rechercher... ». Le tableau
  se filtre sur les demandes pertinentes, dans l'ordre de pertinence, avec un badge « Résultats IA »,
  un score et la justification en infobulle. Modifier le texte ou « Revenir à la recherche classique »
  quitte le mode IA. La recherche texte reste le comportement par défaut et le repli en cas
  d'indisponibilité, de délai de 8 s dépassé ou de limite atteinte.
- **Demandes similaires** : au clic sur « Soumettre » d'une nouvelle demande, l'IA cherche des demandes
  passées ou en cours sur le même sujet. S'il y en a, une fenêtre les affiche avec la réponse apportée
  par la Compliance ; l'utilisateur choisit « Soumettre quand même » ou « Annuler et revenir au
  formulaire ». Sans demande similaire, aucune fenêtre ; en cas d'échec, la soumission continue. Les
  références proposées et le choix sont enregistrés dans la colonne `DemandesSimilaires` et visibles
  de la Compliance sur la demande et dans l'email de nouvelle demande.
  Pour chaque demande proposée, l'IA donne « Ce qu'il faut retenir » (résumé signalé comme tel) et
  désigne le message décisif de la Compliance, affiché mot pour mot, l'échange complet restant
  dépliable. Une demande encore ouverte sur le même sujet est signalée « Même demande en cours ».
  Ne sont jamais proposées : les demandes annulées ou rejetées pour doublon, les demandes terminées
  sans réponse de la Compliance, et celles que l'IA juge sans intérêt. Les types « Validation
  documentation commerciale » et « Franchissement de seuil » ne déclenchent pas l'analyse.
- **Qui voit quoi** : la Compliance voit toutes les demandes. Un collaborateur ne voit que les siennes
  et celles où il est en copie ; à la création, il voit en plus les demandes d'autres personnes que la
  Compliance a cochées « Partageable » (case du panneau « Partage » de la demande), sans nom ni
  département du demandeur et sans lien vers la demande. Ce filtrage est fait dans le navigateur,
  comme celui du tableau : la protection réelle dépend des permissions SharePoint.
- **Réponse apportée** : commentaires publics des membres de la Compliance, plus le motif d'un rejet.
  Les notes internes (`[INTERNE]`) ne sont jamais envoyées au modèle (double filtre navigateur et
  serveur). Le texte affiché vient toujours des données, jamais du modèle ; le serveur ignore toute
  référence absente du catalogue envoyé.
- **Catalogue** : construit par le navigateur (référence, objet, type, statut, extraits de la description
  et de la réponse, et demandeur et département quand l'utilisateur y a droit), envoyé avec la
  requête. Le catalogue est placé en tête du prompt pour profiter du cache de prompt automatique d'OpenAI.
- **Prompts** : `api/shared/prompts-tracker-ai.md` (sections RECHERCHE, SIMILAIRES et GLOSSAIRE, modifiables
  sans toucher au code).

Configuration : aucune variable supplémentaire n'est obligatoire. Les fonctions utilisent `OPENAI_API_KEY`
et `OPENAI_BASE_URL` de Doc Review, et par défaut le même modèle (`OPENAI_MODEL_SYNTHESE`, sinon le modèle
par défaut de Doc Review) avec un effort de raisonnement `low`. Réglages facultatifs (Application
Settings des Functions, jamais dans le dépôt) :

| Variable | Rôle |
|---|---|
| `OPENAI_MODEL_RECHERCHE` | Modèle de la recherche IA (défaut : `OPENAI_MODEL_SYNTHESE`) |
| `OPENAI_MODEL_SIMILAIRES` | Modèle des demandes similaires (défaut : celui de la recherche) |
| `OPENAI_EFFORT_RECHERCHE` | Effort de raisonnement de la recherche (défaut `low`) |
| `OPENAI_EFFORT_SIMILAIRES` | Effort de raisonnement des demandes similaires (défaut : celui de la recherche) |
| `OPENAI_TIMEOUT_MS_TRACKER` | Délai maximal d'un appel au modèle en ms (défaut 7000, de 1000 à 40000) |

Les appels sont synchrones (sans mode background) : un modèle rapide ou un effort faible est nécessaire pour
tenir le délai de 8 s du navigateur.

Colonnes à créer dans la liste « Tracker CPL Demandes » :

| Colonne (nom interne) | Type | Rôle |
|---|---|---|
| `DemandesSimilaires` | Plusieurs lignes de texte (texte brut) | Demandes proposées et choix de l'utilisateur (JSON) |
| `Partageable` | Oui/Non | Demande visible, anonymisée, dans les demandes similaires des autres |

Sans ces colonnes, les fonctions marchent mais l'enregistrement du choix et le partage échouent (message
d'erreur pour le partage).
