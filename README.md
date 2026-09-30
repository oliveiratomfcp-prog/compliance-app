# Compliance Portal — Eiffel Investment Group

Portail statique (Azure Static Web App) avec des Azure Functions (`api/`) qui font le pont vers
Microsoft Graph / SharePoint et vers des API tierces, en gardant les secrets côté serveur.

## Configuration — variables d'environnement des Azure Functions

Ces variables se définissent dans les **Application Settings** de l'Azure Static Web App / Function
App, depuis le Portail Azure. Elles ne doivent **jamais** être commitées dans le dépôt (ni dans le
front-end, ni dans `api/local.settings.json`).

| Variable | Utilisée par | Rôle |
|---|---|---|
| `OPENAI_API_KEY` | `api/analyze` | Clé de l'API OpenAI (Doc Review). |

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
